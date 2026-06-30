<#
.SYNOPSIS
    End-to-end setup for the Adaptive Card Agent AI teammate.

.DESCRIPTION
    Runs all phases in order:
      1. Create single-tenant Entra app registration + service principal
         (adds redirect URI + wids claim required by a365 CLI)
      2. Deploy Azure Bot Service + Teams channel via Bicep (rg-aacebo)
      3. Update a365.config.json with the new clientAppId
      4. Run a365 setup all --m365
         (pauses for admin consent URL — script prints it and waits)
      5. Patch OAuth2 grant leading-space bug (known a365 CLI defect)
      6. Update .env with new blueprint credentials
      7. Update manifest files + re-zip (manifest/agent/ and manifest/bot/)
      8. Print upload instructions

    Safe to re-run. Skips Entra app creation if one with the same display name
    already exists (pass -Force to rotate the secret).

.PARAMETER TunnelEndpoint
    Devtunnel messaging URL. Default: https://aacebo-3978.use.devtunnels.ms/api/messages

.PARAMETER AppDisplayName
    Display name for the Entra app + Bot Service. Default: Adaptive Card Agent

.PARAMETER ResourceGroup
    Azure resource group. Default: rg-aacebo

.PARAMETER Force
    Rotate the Entra app client secret even if the app already exists.

.PARAMETER SkipBicep
    Skip the Azure Bot Service Bicep deployment (use when bot service already exists).

.PARAMETER SkipA365Setup
    Skip a365 setup all (use when blueprint already exists and you only want to fix
    grants / env / manifest).

.EXAMPLE
    pwsh -File infra/setup.ps1
    pwsh -File infra/setup.ps1 -Force
    pwsh -File infra/setup.ps1 -SkipBicep -SkipA365Setup
#>
[CmdletBinding()]
param(
    [string]$TunnelEndpoint   = 'https://aacebo-3978.use.devtunnels.ms/api/messages',
    [string]$AppDisplayName   = 'Adaptive Card Agent',
    [string]$ResourceGroup    = 'aacebo-rg',
    [switch]$Force,
    [switch]$SkipBicep,
    [switch]$SkipA365Setup
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = (Get-Item "$PSScriptRoot/..").FullName

function Write-Step([string]$msg) { Write-Host "`n==> $msg" -ForegroundColor Cyan }
function Write-Ok([string]$msg)   { Write-Host "  OK  $msg" -ForegroundColor Green }
function Write-Warn([string]$msg) { Write-Host "  WARN $msg" -ForegroundColor Yellow }
function Write-Fail([string]$msg) { Write-Host "  FAIL $msg" -ForegroundColor Red; exit 1 }

# ---------------------------------------------------------------------------
# Preflight
# ---------------------------------------------------------------------------
Write-Step 'Checking az CLI session'
$account = az account show --output json 2>$null | ConvertFrom-Json
if (-not $account) { Write-Fail 'az CLI not signed in. Run: az login' }
Write-Ok "Signed in as $($account.user.name)"
Write-Ok "Tenant: $($account.tenantId)"
Write-Ok "Subscription: $($account.name) ($($account.id))"

$tenantId = $account.tenantId

# ---------------------------------------------------------------------------
# Phase 1 — Entra app registration (single-tenant)
# ---------------------------------------------------------------------------
Write-Step "Phase 1 — Entra app registration ($AppDisplayName)"

$existing = az ad app list --display-name $AppDisplayName --output json | ConvertFrom-Json
$botAppId     = $null
$clientSecret = $null

if ($existing -and $existing.Length -gt 0 -and -not $Force) {
    $botAppId = $existing[0].appId
    Write-Ok "Reusing existing app: $botAppId (pass -Force to rotate secret)"

    $secretJson = az ad app credential reset `
        --id $botAppId `
        --append `
        --display-name "aid-teams-$(Get-Date -Format 'yyyyMMdd')" `
        --years 1 `
        --output json | ConvertFrom-Json
    $clientSecret = $secretJson.password
    Write-Ok 'Client secret rotated'
} else {
    if ($existing -and $existing.Length -gt 0) {
        $botAppId = $existing[0].appId
        Write-Warn "Existing app found but -Force specified — reusing, rotating secret"
    } else {
        Write-Ok 'Creating new single-tenant app registration'
        $created = az ad app create `
            --display-name $AppDisplayName `
            --sign-in-audience AzureADMyOrg `
            --output json | ConvertFrom-Json
        $botAppId = $created.appId
        Write-Ok "Created: $botAppId"
    }

    # Service principal
    $sp = az ad sp list --filter "appId eq '$botAppId'" --output json | ConvertFrom-Json
    if (-not $sp -or $sp.Length -eq 0) {
        az ad sp create --id $botAppId --output none
        Write-Ok 'Service principal created'
    } else {
        Write-Ok 'Service principal already exists'
    }

    # Add redirect URIs required by a365 CLI
    # - https://entra.microsoft.com/TokenAuthorize  : admin consent flow
    # - ms-appx-web://Microsoft.AAD.BrokerPlugin/.. : Windows Account Manager (WAM) sign-in
    Write-Step 'Adding redirect URIs (required by a365 CLI)'
    az ad app update `
        --id $botAppId `
        --public-client-redirect-uris "https://entra.microsoft.com/TokenAuthorize" "ms-appx-web://Microsoft.AAD.BrokerPlugin/$botAppId" `
        --output none
    Write-Ok "Redirect URIs added"

    # Add wids optional claim (required for Global Admin role detection by a365 CLI)
    Write-Step "Adding 'wids' optional claim to access token"
    az rest --method PATCH `
        --url "https://graph.microsoft.com/v1.0/applications(appId='$botAppId')" `
        --headers 'Content-Type=application/json' `
        --body '{"optionalClaims":{"accessToken":[{"name":"wids","essential":false,"additionalProperties":[]}]}}' `
        --output none
    Write-Ok "'wids' claim added"

    # API permissions for a365 CLI requirements check
    Write-Step 'Configuring API permissions'
    $msGraphAppId  = '00000003-0000-0000-c000-000000000000'
    $requiredScopes = @(
        'User.Read',
        'Application.ReadWrite.All',
        'AgentIdentityBlueprint.ReadWrite.All',
        'AgentIdentityBlueprint.UpdateAuthProperties.All',
        'DelegatedPermissionGrant.ReadWrite.All',
        'Directory.Read.All'
    )
    $graphSp  = az ad sp show --id $msGraphAppId --output json | ConvertFrom-Json
    $scopeMap = @{}
    foreach ($s in $graphSp.oauth2PermissionScopes) { $scopeMap[$s.value] = $s.id }

    foreach ($scopeName in $requiredScopes) {
        if ($scopeMap.ContainsKey($scopeName)) {
            az ad app permission add `
                --id $botAppId `
                --api $msGraphAppId `
                --api-permissions "$($scopeMap[$scopeName])=Scope" `
                --output none 2>$null
            Write-Ok "+ $scopeName"
        } else {
            Write-Warn "Scope not found on Graph SP: $scopeName — add manually in Entra portal"
        }
    }

    # Admin consent
    Write-Step 'Granting admin consent'
    try {
        az ad app permission admin-consent --id $botAppId --output none
        Write-Ok 'Admin consent granted'
    } catch {
        Write-Warn 'Admin consent failed — grant manually: Entra portal > App registrations > API permissions > Grant admin consent'
    }

    # Client secret
    Write-Step 'Generating client secret'
    $secretJson = az ad app credential reset `
        --id $botAppId `
        --append `
        --display-name "aid-teams-$(Get-Date -Format 'yyyyMMdd')" `
        --years 1 `
        --output json | ConvertFrom-Json
    $clientSecret = $secretJson.password
    Write-Ok 'Client secret created'
}

Write-Ok "botAppId: $botAppId"

# Force az re-login so the new wids claim is in the token used by a365 CLI
Write-Step 'Refreshing az token (so wids claim is present for a365 CLI)'
az account get-access-token --output none 2>$null
Write-Ok 'Token refreshed'

# ---------------------------------------------------------------------------
# Phase 2 — Azure Bot Service via Bicep
# ---------------------------------------------------------------------------
if (-not $SkipBicep) {
    Write-Step "Phase 2 — Azure Bot Service (rg: $ResourceGroup)"

    # Ensure resource group exists
    $rg = az group show --name $ResourceGroup --output json 2>$null | ConvertFrom-Json
    if (-not $rg) {
        Write-Ok "Creating resource group $ResourceGroup"
        az group create --name $ResourceGroup --location eastus --output none
        Write-Ok "Resource group created"
    } else {
        Write-Ok "Resource group $ResourceGroup already exists ($($rg.location))"
    }

    # Bot Service must use the blueprint ID as msaAppId so BF routes to the right identity
    $generatedConfigPathForBicep = Join-Path $ProjectRoot 'a365.generated.config.json'
    $bicepBotAppId = $botAppId
    if (Test-Path $generatedConfigPathForBicep) {
        $genCfg = Get-Content $generatedConfigPathForBicep -Raw | ConvertFrom-Json
        if ($genCfg.agentBlueprintId) {
            $bicepBotAppId = $genCfg.agentBlueprintId
            Write-Ok "Using blueprint ID for bot msaAppId: $bicepBotAppId"
        }
    }

    $botName   = $AppDisplayName.ToLower() -replace '[^a-z0-9]', '-'
    $bicepPath = Join-Path $PSScriptRoot 'main.bicep'
    $result = az deployment group create `
        --resource-group $ResourceGroup `
        --template-file $bicepPath `
        --parameters botAppId=$bicepBotAppId messagingEndpoint=$TunnelEndpoint botDisplayName=$AppDisplayName botName=$botName `
        --output json 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Host $result
        Write-Fail "Bicep deployment failed (exit $LASTEXITCODE)"
    }
    Write-Ok "Bot Service deployed to $ResourceGroup"
} else {
    Write-Warn 'Phase 2 skipped (-SkipBicep)'
}

# ---------------------------------------------------------------------------
# Phase 3 — Update a365.config.json
# ---------------------------------------------------------------------------
Write-Step 'Phase 3 — Updating a365.config.json'
$configPath = Join-Path $ProjectRoot 'a365.config.json'
$config = Get-Content $configPath -Raw | ConvertFrom-Json
$config.clientAppId = $botAppId
$config | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8
Write-Ok "clientAppId set to $botAppId"

# ---------------------------------------------------------------------------
# Phase 4 — a365 setup all --m365
# ---------------------------------------------------------------------------
if (-not $SkipA365Setup) {
    Write-Step 'Phase 4 — Running a365 setup all --m365'
    Write-Host ''
    Write-Host '  The CLI will print an admin consent URL.' -ForegroundColor Yellow
    Write-Host '  Open it in your browser, accept as Global Admin, then press ENTER here.' -ForegroundColor Yellow
    Write-Host ''

    Set-Location $ProjectRoot
    $setupJob = Start-Job -ScriptBlock {
        Set-Location $using:ProjectRoot
        'y' | a365 setup all --m365 2>&1
    }

    # Stream output so the consent URL is visible
    $consentUrl = $null
    while ($setupJob.State -eq 'Running') {
        $output = Receive-Job $setupJob
        if ($output) {
            $output | ForEach-Object {
                Write-Host "  $_"
                if ($_ -match 'Consent URL:') { $consentUrl = $_.Trim() }
            }
        }
        Start-Sleep -Milliseconds 500
    }
    # Flush remaining output
    Receive-Job $setupJob | ForEach-Object { Write-Host "  $_" }
    Remove-Job $setupJob

    Write-Host ''
    Write-Host '  If you see a consent URL above, open it in your browser and accept.' -ForegroundColor Yellow
    Read-Host '  Press ENTER once you have accepted the consent'

    # Re-run to pick up the grants now that consent is done
    Write-Step 'Phase 4b — Re-running a365 setup all --m365 to apply grants'
    Set-Location $ProjectRoot
    'y' | a365 setup all --m365 2>&1 | ForEach-Object { Write-Host "  $_" }

    Write-Ok 'a365 setup completed'
} else {
    Write-Warn 'Phase 4 skipped (-SkipA365Setup)'
}

# ---------------------------------------------------------------------------
# Phase 5 — Fix OAuth2 grant leading-space bug
# ---------------------------------------------------------------------------
Write-Step 'Phase 5 — Patching OAuth2 grant leading-space bug'

$generatedConfigPath = Join-Path $ProjectRoot 'a365.generated.config.json'
if (-not (Test-Path $generatedConfigPath)) {
    Write-Warn 'a365.generated.config.json not found — skipping (run setup first)'
} else {
    $generated   = Get-Content $generatedConfigPath -Raw | ConvertFrom-Json
    $blueprintId = $generated.agentBlueprintId

    if (-not $blueprintId) {
        Write-Warn 'agentBlueprintId is empty in generated config — skipping grant patch'
    } else {
        $blueprintSp = az ad sp list --filter "appId eq '$blueprintId'" --query '[0].id' -o tsv 2>$null

        if ($blueprintSp) {
            $grantsJson = az rest --method GET `
                --url "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$blueprintSp'" `
                --output json | ConvertFrom-Json

            $patched = 0
            foreach ($grant in $grantsJson.value) {
                if ($grant.scope -match '^\s') {
                    $cleaned = $grant.scope.TrimStart()
                    az rest --method PATCH `
                        --url "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($grant.id)" `
                        --headers 'Content-Type=application/json' `
                        --body "{`"scope`": `"$cleaned`"}" `
                        --output none
                    $patched++
                }
            }
            Write-Ok "Patched $patched grant(s) for blueprint $blueprintId"
        } else {
            Write-Warn "Blueprint SP not found for $blueprintId — skipping grant patch"
        }
    }
}

# ---------------------------------------------------------------------------
# Phase 6 — Update .env
# ---------------------------------------------------------------------------
Write-Step 'Phase 6 — Updating .env'

if (-not (Test-Path $generatedConfigPath)) {
    Write-Warn 'a365.generated.config.json not found — skipping .env update'
} else {
    $generated     = Get-Content $generatedConfigPath -Raw | ConvertFrom-Json
    $blueprintId   = $generated.agentBlueprintId
    $blueprintSecret = $null

    if (-not $blueprintId) {
        Write-Warn 'agentBlueprintId is empty — .env will use botAppId + clientSecret from Phase 1'
        $blueprintId     = $botAppId
        $blueprintSecret = $clientSecret
    } elseif ($generated.agentBlueprintClientSecretProtected -eq $true) {
        try {
            Add-Type -AssemblyName System.Security
            $blueprintSecret = [System.Text.Encoding]::UTF8.GetString(
                [System.Security.Cryptography.ProtectedData]::Unprotect(
                    [System.Convert]::FromBase64String($generated.agentBlueprintClientSecret),
                    $null,
                    [System.Security.Cryptography.DataProtectionScope]::CurrentUser))
        } catch {
            Write-Warn 'Could not decrypt DPAPI secret — using clientSecret from Phase 1'
            $blueprintSecret = $clientSecret
        }
    } else {
        $blueprintSecret = $generated.agentBlueprintClientSecret
    }

    $envPath = Join-Path $ProjectRoot '.env'
    $envContent = @"
connections__service_connection__settings__clientId=$blueprintId
agent_id=$blueprintId
connections__service_connection__settings__clientSecret=$blueprintSecret
connections__service_connection__settings__tenantId=$tenantId
connections__service_connection__settings__scopes=5a807f24-c9de-44ee-a3a7-329e88a00ffc/.default
connections__obo_connection__type=OAuthConnection
connections__obo_connection__settings__clientId=$blueprintId
connections__obo_connection__settings__clientSecret=$blueprintSecret
connections__obo_connection__settings__tenantId=$tenantId
connections__obo_connection__settings__scopes=https://graph.microsoft.com/.default
connectionsMap__0__serviceUrl=*
connectionsMap__0__connection=service_connection
agentic_altBlueprintConnectionName=service_connection
agentic_scopes=https://graph.microsoft.com/.default
agentic_connectionName=AgenticAuthConnection
ENABLE_A365_OBSERVABILITY_EXPORTER=false
agent365Observability__agentId=$blueprintId
agent365Observability__agentName=$AppDisplayName
agent365Observability__agentDescription=Your #1 Super Agent
agent365Observability__tenantId=$tenantId
agent365Observability__agentBlueprintId=$blueprintId
agent365Observability__clientId=$blueprintId
agent365Observability__clientSecret=$blueprintSecret
"@
    Set-Content $envPath $envContent -Encoding UTF8
    Write-Ok '.env updated'
}

# ---------------------------------------------------------------------------
# Phase 7 — Update manifest + re-zip
# ---------------------------------------------------------------------------
Write-Step 'Phase 7 — Updating manifest files and re-zipping'

if (-not (Test-Path $generatedConfigPath)) {
    Write-Warn 'a365.generated.config.json not found — skipping manifest update'
} else {
    $generated   = Get-Content $generatedConfigPath -Raw | ConvertFrom-Json
    $blueprintId = $generated.agentBlueprintId

    if (-not $blueprintId) {
        Write-Warn 'agentBlueprintId is empty — skipping manifest update'
    } else {
        $manifestDir   = Join-Path $ProjectRoot 'manifest'
        $agentDir      = Join-Path $manifestDir 'agent'
        $botDir        = Join-Path $manifestDir 'bot'
        $agentManifest = Join-Path $agentDir 'manifest.json'
        $agenticPath   = Join-Path $agentDir 'agenticUserTemplateManifest.json'
        $botManifest   = Join-Path $botDir 'manifest.json'
        $agentZip      = Join-Path $agentDir 'manifest.zip'
        $botZip        = Join-Path $botDir 'manifest.zip'

        $existing   = Get-Content $agentManifest -Raw | ConvertFrom-Json
        $vParts     = ($existing.version -replace '[^\d\.]','') -split '\.'
        $newVersion = "$($vParts[0]).$($vParts[1]).$([int]$vParts[2] + 1)"

        $developer = [ordered]@{
            name          = 'Microsoft Corporation'
            mpnId         = ''
            websiteUrl    = 'https://go.microsoft.com/fwlink/?LinkId=518028'
            privacyUrl    = 'https://go.microsoft.com/fwlink/?LinkId=518028'
            termsOfUseUrl = 'https://shares.datatransfer.microsoft.com/assets/Microsoft_Terms_of_Use.html'
        }

        # Agent identity manifest — personal scope only, no RSC/webApplicationInfo
        $agentManifestObj = [ordered]@{
            '$schema'            = 'https://developer.microsoft.com/en-us/json-schemas/teams/vdevPreview/MicrosoftTeams.schema.json'
            id                   = $blueprintId
            version              = $newVersion
            manifestVersion      = 'devPreview'
            accentColor          = '#9ec9d9'
            name                 = [ordered]@{ short = $AppDisplayName; full = $AppDisplayName }
            description          = [ordered]@{ short = 'Your #1 Super Agent'; full = 'Your #1 Super Agent' }
            icons                = [ordered]@{ outline = 'outline.png'; color = 'color.png' }
            developer            = $developer
            bots                 = @([ordered]@{
                botId              = $blueprintId
                scopes             = @('personal')
                supportsFiles      = $false
                isNotificationOnly = $false
            })
            agenticUserTemplates = @([ordered]@{
                id   = '7b0926a6-c4ee-445a-a913-bd054594bd09'
                file = 'agenticUserTemplateManifest.json'
            })
            copilotAgents        = [ordered]@{
                customEngineAgents = @([ordered]@{
                    id                    = $blueprintId
                    type                  = 'bot'
                    functionsAs           = 'agenticUserOnly'
                    agenticUserTemplateId = '7b0926a6-c4ee-445a-a913-bd054594bd09'
                })
            }
        }
        $agentManifestObj | ConvertTo-Json -Depth 10 | Set-Content $agentManifest -Encoding UTF8
        Write-Ok "agent/manifest.json updated (id: $blueprintId, version: $newVersion)"

        $agentic = Get-Content $agenticPath -Raw | ConvertFrom-Json
        $agentic.agentIdentityBlueprintId = $blueprintId
        $agentic | ConvertTo-Json -Depth 10 | Set-Content $agenticPath -Encoding UTF8
        Write-Ok 'agent/agenticUserTemplateManifest.json updated'

        Remove-Item $agentZip -Force -ErrorAction SilentlyContinue
        Compress-Archive `
            -Path "$agentDir/manifest.json", "$agentDir/agenticUserTemplateManifest.json", "$agentDir/color.png", "$agentDir/outline.png" `
            -DestinationPath $agentZip
        Write-Ok "agent/manifest.zip created ($($(Get-Item $agentZip).Length) bytes)"

        # Bot manifest — personal/team/groupChat scopes, RSC, webApplicationInfo
        $botManifestObj = [ordered]@{
            '$schema'          = 'https://developer.microsoft.com/en-us/json-schemas/teams/vdevPreview/MicrosoftTeams.schema.json'
            id                 = $blueprintId
            version            = $newVersion
            manifestVersion    = 'devPreview'
            accentColor        = '#9ec9d9'
            name               = [ordered]@{ short = $AppDisplayName; full = $AppDisplayName }
            description        = [ordered]@{ short = 'Your #1 Super Agent'; full = 'Your #1 Super Agent' }
            icons              = [ordered]@{ outline = 'outline.png'; color = 'color.png' }
            developer          = $developer
            bots               = @([ordered]@{
                botId              = $blueprintId
                scopes             = @('personal', 'team', 'groupChat')
                supportsFiles      = $false
                isNotificationOnly = $false
            })
            webApplicationInfo = [ordered]@{
                id       = $blueprintId
                resource = 'https://RscBasedStoreApp'
            }
            authorization      = [ordered]@{
                permissions = [ordered]@{
                    resourceSpecific = @(
                        [ordered]@{ name = 'ChannelMessage.Read.Group'; type = 'Application' },
                        [ordered]@{ name = 'ChatMessage.Read.Chat';     type = 'Application' }
                    )
                }
            }
            permissions              = @('identity', 'messageTeamMembers')
            supportsChannelFeatures  = 'tier1'
            validDomains             = @()
        }
        $botManifestObj | ConvertTo-Json -Depth 10 | Set-Content $botManifest -Encoding UTF8
        Write-Ok "bot/manifest.json updated (id: $blueprintId, version: $newVersion)"

        Remove-Item $botZip -Force -ErrorAction SilentlyContinue
        Compress-Archive `
            -Path "$botDir/manifest.json", "$botDir/color.png", "$botDir/outline.png" `
            -DestinationPath $botZip
        Write-Ok "bot/manifest.zip created ($($(Get-Item $botZip).Length) bytes)"
    }
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host '========================================' -ForegroundColor Green
Write-Host '  Setup complete!' -ForegroundColor Green
Write-Host '========================================' -ForegroundColor Green
Write-Host ''
Write-Host 'Next steps — upload both manifests:' -ForegroundColor Cyan
Write-Host ''
Write-Host '  Agent identity (admin.microsoft.com > Agents > Upload custom agent):' -ForegroundColor White
Write-Host "    $ProjectRoot\manifest\agent\manifest.zip" -ForegroundColor White
Write-Host ''
Write-Host '  Teams bot (Teams > Apps > Manage your apps > Upload an app):' -ForegroundColor White
Write-Host "    $ProjectRoot\manifest\bot\manifest.zip" -ForegroundColor White
Write-Host ''
Write-Host 'Then start the dev server:' -ForegroundColor Cyan
Write-Host '  npm run dev' -ForegroundColor White
Write-Host ''
