@description('The blueprint app ID (ed26402d-...) — must match botId in manifest')
param botAppId string
param messagingEndpoint string
param botDisplayName string = 'Adaptive Card Agent'
param botName string = 'adaptive-card-agent'
param tenantId string = tenant().tenantId

resource botService 'Microsoft.BotService/botServices@2022-09-15' = {
  name: botName
  location: 'global'
  sku: { name: 'F0' }
  kind: 'azurebot'
  properties: {
    displayName: botDisplayName
    msaAppId: botAppId
    msaAppType: 'MultiTenant'
    msaAppTenantId: ''
    endpoint: messagingEndpoint
    schemaTransformationVersion: '1.3'
  }
}

resource teamsChannel 'Microsoft.BotService/botServices/channels@2022-09-15' = {
  parent: botService
  name: 'MsTeamsChannel'
  location: 'global'
  properties: {
    channelName: 'MsTeamsChannel'
    properties: {
      isEnabled: true
    }
  }
}
