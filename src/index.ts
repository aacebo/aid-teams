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
  // const agentRef = { ...ctx.ref, serviceUrl: 'https://smba.trafficmanager.net/teams' };

  // Pass a mutated ctx to next() so the router uses it as mergedContext instead of rebuilding from toInterface()
  return ctx.next({
    ...ctx,
    send: (activity, conversationRef) => sender.send(toActivityParams(activity), conversationRef ?? ctx.ref),
  });
});

// Teams — standard Bot Framework channel (S2S auth = receives all group chat messages without @mention)
app.on('message', async (ctx) => {
  if (ctx.activity.channelId !== 'msteams') { await ctx.next(ctx); return; }

  const convType = ctx.activity.conversation.conversationType; // 'personal' | 'groupChat' | 'channel'
  ctx.log.info(`teams [${convType}] from ${ctx.activity.from.name}: ${ctx.activity.text}`);
  await ctx.send(`you said: "${ctx.activity.text}"`);
});

// Agent 365 — all M365 surface notifications arrive on channelId "agents"
// Email: entities[].type === "emailNotification"
// Word/Excel/PowerPoint: channelData.productContext
app.on('message', async (ctx) => {
  if (ctx.activity.channelId !== 'agents') { await ctx.next(ctx); return; }

  const channelData = ctx.activity.channelData as AgentsChannelData;
  const productContext = channelData?.productContext;
  const isEmail = ctx.activity.entities?.some(
    (e: any) => e.type === 'emailNotification'
  );

  if (!isEmail && !productContext) return;

  if (isEmail) {
    ctx.log.info(`email from ${ctx.activity.from.id}: ${ctx.activity.text}`);
  } else {
    ctx.log.info(`${productContext} comment from ${ctx.activity.from.name}: ${ctx.activity.text}`);
  }

  await ctx.send({
    channelId: 'agents',
    type: 'message',
    text: `you said: "${ctx.activity.text}"`,
    recipient: ctx.activity.from,
    from: ctx.activity.recipient,
    channelData,
  });
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