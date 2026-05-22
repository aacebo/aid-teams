export type AgenticRecipient = {
  readonly id: string;
  readonly name: string;
  readonly role: string;
  readonly tenantId?: string;
  readonly agenticUserId: string;
  readonly agenticAppId: string;
  readonly agenticAppBlueprintId: string;
  readonly callbackUri?: string;
};
