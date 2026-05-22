import { App } from '@microsoft/teams.apps';
import { ActivitySender } from '@microsoft/teams.apps/dist/activity-sender';
import { DevtoolsPlugin } from '@microsoft/teams.dev';
import { Client as HttpClient } from '@microsoft/teams.common';
import { AgentTokenClient } from './agent-token-client';
import type { AgentsChannelData } from './models/index';

const tokenClient = new AgentTokenClient({
  tenantId: process.env.TENANT_ID!,
  clientId: process.env.CLIENT_ID!,
  clientSecret: process.env.CLIENT_SECRET!,
  agentIdentityId: process.env.AGENT_IDENTITY_ID!,
  agentUserOid: process.env.AGENT_USER_OID!,
});

const app = new App({
  managedIdentityClientId: process.env.MANAGED_IDENTITY_PRINCIPAL_ID,
  plugins: [new DevtoolsPlugin()],
});

app.event('error', ({ error, activity }) => {
  app.log.error(`${error.message}\n${error.stack ?? ''}`, { activity });
});

app.use(async (ctx) => {
  if (ctx.activity.channelId !== 'agents') return ctx.next();

  const agentToken = await tokenClient.getToken('https://api.botframework.com/.default');
  const http = new HttpClient({ token: `Bearer ${agentToken}` });
  (ctx as any).activitySender = new ActivitySender(http, app.log);

  return ctx.next();
});

// Teams — standard Bot Framework channel (S2S auth = receives all group chat messages without @mention)
app.on('message', async ({ activity, send, next, log }) => {
  if (activity.channelId !== 'msteams') { await next(); return; }

  const convType = activity.conversation.conversationType; // 'personal' | 'groupChat' | 'channel'
  log.info(`teams [${convType}] from ${activity.from.name}: ${activity.text}`);
  await send(`you said: "${activity.text}"`);
});

// Agent 365 — all M365 surface notifications arrive on channelId "agents"
// Email: entities[].type === "emailNotification"
// Word/Excel/PowerPoint: channelData.productContext
app.on('message', async ({ activity, send, next, log }) => {
  if (activity.channelId !== 'agents') { await next(); return; }

  const channelData = activity.channelData as AgentsChannelData;
  const productContext = channelData?.productContext;
  const isEmail = activity.entities?.some(
    (e) => (e as unknown as { type: string }).type === 'emailNotification'
  );

  if (!isEmail && !productContext) return;

  if (isEmail) {
    log.info(`email from ${activity.from.id}: ${activity.text}`);
  } else {
    log.info(`${productContext} comment from ${activity.from.name}: ${activity.text}`);
  }

  await send(`you said: "${activity.text}"`);
});

(async () => {
  await app.start(process.env.PORT || 3978).catch((err) => app.log.error(err));

  const shutdown = async () => {
    await app.stop();
    process.exit(0);
  };

  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
})();