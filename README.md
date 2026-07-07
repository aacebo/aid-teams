# aid-teams

Agent Identity Teams app — a bot that receives messages from Teams, email notifications, and Microsoft 365 product contexts (Word, Excel, PowerPoint) via the `agents` channel.

---

## Prerequisites

| Tool | Install |
|---|---|
| Node.js 22+ | https://nodejs.org |
| Azure CLI | https://learn.microsoft.com/cli/azure/install-azure-cli |
| PowerShell 7+ | https://learn.microsoft.com/powershell/scripting/install/installing-powershell |
| a365 CLI | `npm install -g @microsoft/a365@latest` |
| Dev tunnel | https://learn.microsoft.com/azure/developer/dev-tunnels/get-started |

You also need:
- An Azure subscription with permission to create resource groups and Bot Services
- A Microsoft 365 tenant where you have Global Admin (or enough delegated admin) access for Entra app registration and consent

---

## 1. Clone and install

```bash
git clone https://github.com/aacebo/aid-teams.git
cd aid-teams
npm install
```

---

## 2. Configure a365.config.json

Fill in `a365.config.json` before running setup:

```json
{
  "tenantId": "<your Azure AD tenant ID>",
  "subscriptionId": "<your Azure subscription ID>",
  "clientAppId": "",
  "resourceGroup": "<resource group name to create or reuse>",
  "agentIdentityDisplayName": "<display name for the agent identity>",
  "agentBlueprintDisplayName": "<display name shown in Teams>",
  "agentUserDisplayName": "<short name for the agent user>",
  "agentUserPrincipalName": "<UPN for the agent user, e.g. myagent@contoso.onmicrosoft.com>",
  "managerEmail": "<email of the user who will be the agent's manager>",
  "agentUserUsageLocation": "US",
  "agentDescription": "<short description of what the agent does>",
  "messagingEndpoint": "<your dev tunnel URL>/api/messages",
  "deploymentProjectPath": ".",
  "needDeployment": true,
  "environment": "prod",
  "authMode": "both",
  "customBlueprintPermissions": []
}
```

Leave `clientAppId` blank — setup fills it in.

---

## 3. Start a dev tunnel

The bot needs a public HTTPS endpoint. Start a dev tunnel on port 3978:

```bash
devtunnel host -p 3978
```

Copy the tunnel URL (e.g. `https://abc123.use.devtunnels.ms`) and set `messagingEndpoint` in `a365.config.json` to `<tunnel-url>/api/messages`. You can also pass it directly to the setup script (see next step).

---

## 4. Sign in to Azure and run setup

```bash
az login
```

Then run the setup script. It handles all 8 phases automatically:

1. Creates a single-tenant Entra app registration
2. Deploys Azure Bot Service + Teams channel via Bicep
3. Updates `a365.config.json` with the new `clientAppId`
4. Runs `a365 setup all --m365` (pauses and prints an admin consent URL — open it in a browser and accept as Global Admin)
5. Patches a known a365 CLI OAuth2 grant bug
6. Generates `.env` with blueprint credentials
7. Updates and repacks `manifest/agent/manifest.zip` (agent identity) and `manifest/bot/manifest.zip` (Teams bot)
8. Prints upload instructions

```powershell
pwsh -File infra/setup.ps1 -TunnelEndpoint 'https://abc123.use.devtunnels.ms/api/messages'
```

**Flags:**

| Flag | Description |
|---|---|
| `-TunnelEndpoint` | Messaging endpoint override. If omitted, setup resolves the Dev Tunnel `portUri` and appends `/api/messages` |
| `-TunnelId` | Dev Tunnel ID to resolve when `-TunnelEndpoint` is omitted. Default: `aacebo-3978` |
| `-TunnelPort` | Dev Tunnel local port to resolve when `-TunnelEndpoint` is omitted. Default: `3978` |
| `-AppDisplayName` | Entra app + Bot Service display name. Default: `Adaptive Card Agent` |
| `-ResourceGroup` | Azure resource group. Default: `aacebo-rg` |
| `-Force` | Rotate the Entra client secret even if the app already exists |
| `-SkipBicep` | Skip Azure Bot Service deployment (reuse existing) |
| `-SkipA365Setup` | Skip `a365 setup all` (reuse existing blueprint, only fix grants/env/manifest) |

---

## 5. Environment variables

Setup generates a `.env` file. The app also reads these at runtime (add them to `.env` if missing):

| Variable | Description |
|---|---|
| `TENANT_ID` | Azure AD tenant ID |
| `CLIENT_ID` | Entra app registration client ID |
| `CLIENT_SECRET` | Entra app client secret |
| `AGENT_IDENTITY_ID` | Agent identity app/object ID |
| `AGENT_USER_OID` | OID of the agent user |
| `MANAGED_IDENTITY_PRINCIPAL_ID` | Managed identity client ID (optional; used instead of client secret) |
| `PORT` | HTTP port. Default: `3978` |

---

## 6. Run locally

```bash
npm run dev
```

The app starts on `http://localhost:3978` and watches for file changes. Incoming activities are routed through the dev tunnel to this process.

---

## 7. Upload the manifests

After setup completes, upload both packages to your tenant:

**Agent identity** (enables hiring the agent as an AI teammate):
1. Go to the [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Navigate to **Agents → All agents → Upload custom agent**
3. Upload `manifest/agent/manifest.zip`
4. Approve the agent identity when prompted

**Teams bot** (enables installation in teams, channels, and group chats):
1. Open Teams → **Apps → Manage your apps → Upload an app**
2. Upload `manifest/bot/manifest.zip`

---

## 8. Deploy and publish

Update the messaging endpoint after switching from dev tunnel to a production URL:

```bash
npm run endpoint
```

Deploy agent configuration without redeploying Azure infrastructure:

```bash
npm run deploy
```

Publish the app package:

```bash
npm run package
```

---

## npm scripts

| Script | Command | Description |
|---|---|---|
| `dev` | `tsx watch -r dotenv/config src/index.ts` | Start with hot reload |
| `start` | `tsx -r dotenv/config src/index.ts` | Start without watching |
| `build` | `tsc --noEmit` | Type-check only |
| `deploy` | `a365 setup all --skip-infrastructure` | Deploy agent config |
| `endpoint` | `a365 setup blueprint --m365 --update-endpoint` | Update messaging endpoint |
| `package` | `a365 publish` | Package and publish the app |
