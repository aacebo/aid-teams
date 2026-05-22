import type { TenantInfo } from '@microsoft/teams.api';

export type AgentsChannelData = {
  readonly tenant?: TenantInfo;
  readonly productContext?: string;
};
