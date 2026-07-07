import { Client as GraphClient } from '@microsoft/teams.graph';
import * as graphEndpoints from '@microsoft/teams.graph-endpoints';
import type { IEndpoints as CalendarEndpoints } from '@microsoft/teams.graph-endpoints/me/calendar';
import { AgentOboTokenClient } from '../agent-obo-token-client';
import { formatClaimsStatus } from './format-claims-status';
import { getCalendarToken } from './get-calendar-token';
import type { CalendarContext } from './types';

type CalendarResponse = CalendarEndpoints['GET /me/calendar']['response'];

export async function runCalendarReadProbe(
  oboTokenClient: AgentOboTokenClient,
  calendarContext: CalendarContext
): Promise<string> {
  const { token, claims } = await getCalendarToken(oboTokenClient, calendarContext);
  const calendar: CalendarResponse = await new GraphClient({ token }).call(graphEndpoints.me.calendar.get, {
    $select: ['id', 'name', 'owner', 'canEdit'],
  });
  const calendarName = calendar.name ?? calendar.id ?? 'default calendar';
  const owner = calendar.owner?.address ?? calendar.owner?.name ?? 'unknown';
  const canEdit = calendar.canEdit === undefined ? 'unknown' : calendar.canEdit ? 'yes' : 'no';

  return [
    'Calendar read succeeded.',
    `Sender: ${calendarContext.senderName ?? calendarContext.userAssertionClaims.preferred_username ?? calendarContext.userAssertionClaims.oid ?? 'unknown'}`,
    `Incoming assertion source: ${calendarContext.userAssertionSource}`,
    formatClaimsStatus('Incoming assertion', calendarContext.userAssertionClaims),
    formatClaimsStatus('Graph token', claims),
    `Calendar: ${calendarName}`,
    `Owner: ${owner}`,
    `Can edit: ${canEdit}`,
  ].join('\n');
}
