import { App } from '@microsoft/teams.apps';
import { ActivitySender } from '@microsoft/teams.apps/dist/activity-sender';
import { DevtoolsPlugin } from '@microsoft/teams.dev';
import { Client as HttpClient } from '@microsoft/teams.common';
import { toActivityParams, MessageActivity } from '@microsoft/teams.api';
import { AgentTokenClient } from './agent-token-client';
import { createPollCard } from './cards/index';
import type { AgenticRecipient } from './models/index';

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

// !!NOTE!!
// this will not be needed once the Teams SDK has official
// Agent Identity support.
app.use(async (ctx) => {
  const recipient = ctx.activity.recipient as AgenticRecipient;
  const agentIdentityId = recipient?.agenticAppId;
  const agentUserOid = recipient?.agenticUserId;

  if (!agentIdentityId || !agentUserOid) return ctx.next();

  const agentToken = await tokenClient.getToken('https://botapi.skype.com/.default', agentIdentityId, agentUserOid);
  const http = new HttpClient({ token: agentToken });
  const sender = new ActivitySender(http, ctx.log);

  return ctx.next({
    ...ctx,
    send: (activity, conversationRef) => {
      let req = toActivityParams(activity);
      req.channelId = ctx.activity.channelId;
      req.recipient = ctx.activity.from;
      req.from = ctx.activity.recipient;
      req.channelData = ctx.activity.channelData;
      ctx.log.debug(`sending activity => ${JSON.stringify(req, null, 2)}`);
      return sender.send(toActivityParams(req), conversationRef ?? ctx.ref);
    },
  });
});

// Teams — standard Bot Framework channel (S2S auth = receives all group chat messages without @mention)
app.on('message', async (ctx) => {
  if (ctx.activity.channelId !== 'msteams') { await ctx.next(ctx); return; }

  await ctx.send(new MessageActivity().addCard('adaptive', createPollCard(
    `You said: "${ctx.activity.text}" — what do you think?`,
    [
      { title: 'Agree', value: 'agree' },
      { title: 'Disagree', value: 'disagree' },
      { title: 'Not sure', value: 'not_sure' },
    ]
  )));
});

// Agent 365 — all M365 surface notifications arrive on channelId "agents"
// Email: entities[].type === "emailNotification"
// Word/Excel/PowerPoint: channelData.productContext
app.on('message', async (ctx) => {
  if (ctx.activity.channelId !== 'agents') { await ctx.next(ctx); return; }
  let card = createPollCard(
    'what do you think?',
    [
      { title: 'Agree', value: 'agree' },
      { title: 'Disagree', value: 'disagree' },
      { title: 'Not sure', value: 'not_sure' },
    ]
  );

  await ctx.send(new MessageActivity('reply').addCard('adaptive', card));
});

app.on('card.action.poll.submit', async ({ send, activity }) => {
  const data = activity.value?.action?.data;
  await send(`You voted: ${data?.choice}`);

  return {
    statusCode: 200,
    type: 'application/vnd.microsoft.activity.message',
    value: 'Vote recorded',
  };
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