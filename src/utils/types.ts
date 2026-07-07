export type HeaderMap = Record<string, readonly string[]>;

export type AgentContext = {
  readonly agentIdentityId: string;
  readonly agentUserOid: string;
  readonly agentBlueprintId: string;
};

export type CalendarContext = {
  readonly agentIdentityId: string;
  readonly senderName?: string;
  readonly userAssertion: string;
  readonly userAssertionClaims: JwtClaims;
  readonly userAssertionSource: string;
};

export type JwtClaims = {
  readonly aud?: string | readonly string[];
  readonly scp?: string;
  readonly oid?: string;
  readonly tid?: string;
  readonly preferred_username?: string;
  readonly name?: string;
  readonly upn?: string;
};
