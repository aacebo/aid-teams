import { createHash } from 'crypto';

type TokenResponse = {
  access_token: string;
  expires_in: number;
};

type CacheEntry = {
  token: string;
  expiresAtMs: number;
};

export type AgentOboTokenClientOptions = {
  readonly tenantId: string;
  readonly clientId: string;
  readonly clientSecret: string;
};

const JWT_BEARER_GRANT = 'urn:ietf:params:oauth:grant-type:jwt-bearer';

export class AgentOboTokenClient {
  private readonly cache = new Map<string, CacheEntry>();

  constructor(private readonly opts: AgentOboTokenClientOptions) {}

  async getToken(scope: string | string[], agentIdentityId: string, userAssertion: string): Promise<string> {
    const normalized = Array.isArray(scope) ? scope.join(' ') : scope;
    const hashedAssertion = createHash('sha256').update(userAssertion).digest('hex');
    const cacheKey = `${agentIdentityId}:${hashedAssertion}:${normalized}`;
    const cached = this.cache.get(cacheKey);

    if (cached && Date.now() < cached.expiresAtMs - 5 * 60 * 1000) {
      return cached.token;
    }

    const res = await this.postToken({
      grant_type: JWT_BEARER_GRANT,
      client_id: this.opts.clientId,
      client_secret: this.opts.clientSecret,
      assertion: userAssertion,
      requested_token_use: 'on_behalf_of',
      scope: normalized,
    });

    this.cache.set(cacheKey, {
      token: res.access_token,
      expiresAtMs: Date.now() + res.expires_in * 1000,
    });

    return res.access_token;
  }

  private async postToken(params: Record<string, string>): Promise<TokenResponse> {
    const url = `https://login.microsoftonline.com/${this.opts.tenantId}/oauth2/v2.0/token`;
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(params).toString(),
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Token request failed [${res.status}]: ${text}`);
    }

    return res.json() as Promise<TokenResponse>;
  }
}
