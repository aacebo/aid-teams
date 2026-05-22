export type AgentLifecycleManager = {
  readonly displayName: string | null;
  readonly userId: string;
  readonly email: string;
  readonly tenantId: string;
  readonly agenticUserId: string;
  readonly agenticAppInstanceId: string;
  readonly agentIdentityBlueprintId: string;
  readonly eventType: string;
};

export type AgentLifecycleValue = {
  readonly valueType: string;
  readonly expirationDateTime?: string;
  readonly manager: AgentLifecycleManager;
};
