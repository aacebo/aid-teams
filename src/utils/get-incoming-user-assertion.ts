import { RAW_REQUEST_ID_FIELD, requestHeadersById } from '../request-capturing-express-adapter';
import { decodeJwtClaims } from './decode-jwt-claims';
import { formatAudience } from './format-audience';
import { normalizeAudience } from './normalize-audience';
import type { AgentContext, HeaderMap, JwtClaims } from './types';

const JWT_PATTERN = /[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}/g;
const PREFERRED_ASSERTION_HEADERS = [
  'authorization',
  'x-ms-user-token',
  'x-ms-agent-user-token',
  'x-ms-delegated-token',
  'x-ms-client-principal-token',
  'x-agent-user-token',
];

export function getIncomingUserAssertion(activity: unknown, agentContext: AgentContext): {
  token: string;
  claims: JwtClaims;
  source: string;
} {
  let headers: HeaderMap | undefined;

  if (activity && typeof activity === 'object' && !Array.isArray(activity)) {
    const channelData = (activity as Record<string, unknown>).channelData;

    if (channelData && typeof channelData === 'object' && !Array.isArray(channelData)) {
      const requestId = (channelData as Record<string, unknown>)[RAW_REQUEST_ID_FIELD];
      if (typeof requestId === 'string') {
        headers = requestHeadersById.get(requestId);
      }
    }
  }

  if (!headers) {
    throw new Error('Missing incoming delegated user assertion. Request headers were not captured for this activity.');
  }

  const ids = new Set([
    agentContext.agentBlueprintId,
    process.env.agent_id ?? '',
    process.env.connections__service_connection__settings__clientId ?? '',
  ]);
  const expectedAudiences = [...ids].flatMap((id) => id ? [id, `api://${id}`, `api://${id}/access_agent`] : []);
  const expectedAudienceSet = new Set(expectedAudiences.map((audience) => audience.toLowerCase()));
  const candidates: Array<{ token: string; source: string; claims?: JwtClaims; matchesExpectedAudience: boolean }> = [];
  const seenTokens = new Set<string>();
  const seenHeaders = new Set<string>();
  const orderedHeaderNames = [
    ...PREFERRED_ASSERTION_HEADERS,
    ...Object.keys(headers).sort(),
  ];

  for (const headerName of orderedHeaderNames) {
    const normalizedName = headerName.toLowerCase();
    if (seenHeaders.has(normalizedName)) continue;

    seenHeaders.add(normalizedName);

    for (const value of headers[normalizedName] ?? []) {
      for (const token of value.match(JWT_PATTERN) ?? []) {
        if (seenTokens.has(token)) continue;

        seenTokens.add(token);

        const claims = decodeJwtClaims(token);
        candidates.push({
          token,
          source: normalizedName,
          claims,
          matchesExpectedAudience: normalizeAudience(claims?.aud).some((audience) =>
            expectedAudienceSet.has(audience.toLowerCase())
          ),
        });
      }
    }
  }

  const matching = candidates.find((candidate) =>
    candidate.matchesExpectedAudience &&
    !!candidate.claims?.oid &&
    !!candidate.claims.tid
  );

  if (matching?.claims) {
    return {
      token: matching.token,
      claims: matching.claims,
      source: matching.source,
    };
  }

  const expected = expectedAudiences.join(', ');
  const candidateSummary = candidates.length === 0
    ? 'none'
    : candidates.map((candidate) => {
      const claims = candidate.claims;

      return [
        `source=${candidate.source}`,
        `aud=${formatAudience(claims?.aud)}`,
        `oid=${claims?.oid ?? 'none'}`,
        `tid=${claims?.tid ?? 'none'}`,
      ].join(' ');
    }).join('; ');
  const matchingAudience = candidates.find((candidate) => candidate.matchesExpectedAudience);

  if (matchingAudience) {
    throw new Error(`Incoming token audience matched the agent blueprint, but the token is not a delegated human user assertion. Expected oid/tid claims. Candidates: ${candidateSummary}`);
  }

  throw new Error(`Missing incoming delegated user assertion for Agent 365 OBO. Expected a JWT with aud one of [${expected}] and human oid/tid claims. Candidates: ${candidateSummary}`);
}
