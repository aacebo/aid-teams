import { Client as GraphClient } from '@microsoft/teams.graph';
import * as graphEndpoints from '@microsoft/teams.graph-endpoints';
import type { IEndpoints as EventEndpoints } from '@microsoft/teams.graph-endpoints/me/events';
import { AgentOboTokenClient } from '../agent-obo-token-client';
import { formatClaimsStatus } from './format-claims-status';
import { getCalendarToken } from './get-calendar-token';
import type { CalendarContext } from './types';

const PROBE_EVENT_DELAY_MS = 15 * 60 * 1000;
const PROBE_EVENT_DURATION_MS = 5 * 60 * 1000;

type CreateEventBody = EventEndpoints['POST /me/events']['body'];
type CreateEventResponse = EventEndpoints['POST /me/events']['response'];

export async function runCalendarWriteProbe(
  oboTokenClient: AgentOboTokenClient,
  calendarContext: CalendarContext
): Promise<string> {
  const { token, claims } = await getCalendarToken(oboTokenClient, calendarContext);
  const start = new Date(Date.now() + PROBE_EVENT_DELAY_MS);
  const end = new Date(start.getTime() + PROBE_EVENT_DURATION_MS);
  const subject = `Agent delegated calendar probe ${new Date().toISOString()}`;
  const graph = new GraphClient({ token });
  const body: CreateEventBody = {
    '@odata.type': '#microsoft.graph.event',
    subject,
    body: {
      '@odata.type': '#microsoft.graph.itemBody',
      contentType: 'text',
      content: 'Created and deleted automatically by the aid-teams delegated permission probe.',
    },
    start: {
      '@odata.type': '#microsoft.graph.dateTimeTimeZone',
      dateTime: start.toISOString().replace(/\.\d{3}Z$/, ''),
      timeZone: 'UTC',
    },
    end: {
      '@odata.type': '#microsoft.graph.dateTimeTimeZone',
      dateTime: end.toISOString().replace(/\.\d{3}Z$/, ''),
      timeZone: 'UTC',
    },
    isReminderOn: false,
    showAs: 'free',
    attendees: [],
  };
  const event: CreateEventResponse = await graph.call(graphEndpoints.me.events.create, body);

  if (!event.id) {
    throw new Error('Graph created probe event without an id.');
  }

  try {
    await graph.call(graphEndpoints.me.events.del, { 'event-id': event.id });
  } catch (error) {
    throw new Error(`Graph created probe event but delete failed. Event id: ${event.id}. ${error instanceof Error ? error.message : String(error)}`);
  }

  return [
    'Calendar write succeeded.',
    `Sender: ${calendarContext.senderName ?? calendarContext.userAssertionClaims.preferred_username ?? calendarContext.userAssertionClaims.oid ?? 'unknown'}`,
    `Incoming assertion source: ${calendarContext.userAssertionSource}`,
    formatClaimsStatus('Incoming assertion', calendarContext.userAssertionClaims),
    formatClaimsStatus('Graph token', claims),
    `Created and deleted probe event: ${event.subject ?? subject}`,
  ].join('\n');
}
