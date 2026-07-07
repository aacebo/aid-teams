import { AgentOboTokenClient } from '../agent-obo-token-client';
import { decodeJwtClaims } from './decode-jwt-claims';
import type { CalendarContext, JwtClaims } from './types';

const GRAPH_CALENDAR_SCOPE = 'https://graph.microsoft.com/Calendars.ReadWrite';

export async function getCalendarToken(
  oboTokenClient: AgentOboTokenClient,
  calendarContext: CalendarContext
): Promise<{ token: string; claims?: JwtClaims }> {
  const token = await oboTokenClient.getToken(
    GRAPH_CALENDAR_SCOPE,
    calendarContext.agentIdentityId,
    calendarContext.userAssertion
  );
  const claims = decodeJwtClaims(token);

  if (!claims?.scp?.split(/\s+/).includes('Calendars.ReadWrite')) {
    const scopes = claims?.scp?.trim();
    const scopeStatus = !scopes
      ? 'Token did not expose an scp claim.'
      : `Token scp does not include Calendars.ReadWrite. scp=${scopes}`;

    throw new Error(`Graph token lacks Calendars.ReadWrite. ${scopeStatus}`);
  }

  return { token, claims };
}
