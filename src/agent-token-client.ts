type TokenResponse = {
  access_token: string;
  expires_in: number;
};

type CacheEntry = {
  token: string;
  expiresAtMs: number;
};

export type AgentTokenClientOptions = {
  readonly tenantId: string;
  readonly clientId: string;
  readonly clientSecret: string;
};

const FMI_EXCHANGE_SCOPE = 'api://AzureADTokenExchange/.default';
const JWT_BEARER_TYPE = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer';

export class AgentTokenClient {
  private readonly cache = new Map<string, CacheEntry>();

  constructor(private readonly opts: AgentTokenClientOptions) {}

  async getToken(scope: string | string[], agentIdentityId: string, userOid: string): Promise<string> {
    const normalized = this.normalizeScope(scope);
    const cacheKey = `${agentIdentityId}:${userOid}:${normalized}`;
    const cached = this.cache.get(cacheKey);

    if (cached && this.isCacheValid(cached)) {
      return cached.token;
    }

    const t1 = await this.fetchT1(agentIdentityId);
    const t2 = await this.fetchT2(agentIdentityId, t1);
    const res = await this.fetchFinal(normalized, agentIdentityId, userOid, t1, t2);

    this.cache.set(cacheKey, {
      token: res.access_token,
      expiresAtMs: Date.now() + res.expires_in * 1000,
    });

    return res.access_token;
  }

  private normalizeScope(scope: string | string[]): string {
    return Array.isArray(scope) ? scope.join(' ') : scope;
  }

  private isCacheValid(entry: CacheEntry): boolean {
    return Date.now() < entry.expiresAtMs - 5 * 60 * 1000;
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

  private async fetchT1(agentIdentityId: string): Promise<string> {
    const res = await this.postToken({
      grant_type: 'client_credentials',
      client_id: this.opts.clientId,
      client_secret: this.opts.clientSecret,
      scope: FMI_EXCHANGE_SCOPE,
      fmi_path: agentIdentityId,
    });

    return res.access_token;
  }

  private async fetchT2(agentIdentityId: string, t1: string): Promise<string> {
    const res = await this.postToken({
      grant_type: 'client_credentials',
      client_id: agentIdentityId,
      client_assertion_type: JWT_BEARER_TYPE,
      client_assertion: t1,
      scope: FMI_EXCHANGE_SCOPE,
    });

    return res.access_token;
  }

  private async fetchFinal(scope: string, agentIdentityId: string, userOid: string, t1: string, t2: string): Promise<TokenResponse> {
    return this.postToken({
      grant_type: 'user_fic',
      client_id: agentIdentityId,
      client_assertion_type: JWT_BEARER_TYPE,
      client_assertion: t1,
      user_federated_identity_credential: t2,
      user_id: userOid,
      requested_token_use: 'on_behalf_of',
      scope,
    });
  }
}
