import { App } from '@microsoft/teams.apps';
import { ActivitySender } from '@microsoft/teams.apps/dist/activity-sender';
import { DevtoolsPlugin } from '@microsoft/teams.dev';
import { Client as HttpClient } from '@microsoft/teams.common';
import { toActivityParams } from '@microsoft/teams.api';
import { AgentTokenClient } from './agent-token-client';
import type { AgenticRecipient, AgentsChannelData } from './models/index';

const tokenClient = new AgentTokenClient({
  tenantId: process.env.connections__service_connection__settings__tenantId!,
  clientId: process.env.connections__service_connection__settings__clientId!,
  clientSecret: process.env.connections__service_connection__settings__clientSecret!,
});

const app = new App({
  // clientId: process.env.connections__service_connection__settings__clientId,
  // clientSecret: process.env.connections__service_connection__settings__clientSecret,
  // tenantId: process.env.connections__service_connection__settings__tenantId,
  activity: {
    mentions: {
      stripText: true
    }
  },
  plugins: [new DevtoolsPlugin()],
});

app.event('error', ({ error, activity }) => {
  app.log.error(`${error.message}\n${error.stack ?? ''}`, { activity });
});

app.use(async (ctx) => {
  const recipient = ctx.activity.recipient as AgenticRecipient;
  const agentIdentityId = recipient?.agenticAppId;
  const agentUserOid = recipient?.agenticUserId;

  if (!agentIdentityId || !agentUserOid) return ctx.next();

  const agentToken = await tokenClient.getToken('https://botapi.skype.com/.default', agentIdentityId, agentUserOid);
  const http = new HttpClient({ token: agentToken });
  const sender = new ActivitySender(http, ctx.log);
  // Agent tokens must be sent to the global smba endpoint, not the per-tenant S2S connector serviceUrl
  const agentRef = { ...ctx.ref, serviceUrl: 'https://smba.trafficmanager.net/teams' };

  // Pass a mutated ctx to next() so the router uses it as mergedContext instead of rebuilding from toInterface()
  return ctx.next({
    ...ctx,
    send: (activity, conversationRef) => sender.send(toActivityParams(activity), conversationRef ?? agentRef),
  });
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