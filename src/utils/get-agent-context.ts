import type { AgenticRecipient } from '../models/index';
import type { AgentContext } from './types';

export function getAgentContext(recipient: unknown): AgentContext | undefined {
  const agenticRecipient = recipient as Partial<AgenticRecipient> | undefined;
  const agentIdentityId = agenticRecipient?.agenticAppId;
  const agentUserOid = agenticRecipient?.agenticUserId;
  const agentBlueprintId =
    agenticRecipient?.agenticAppBlueprintId ??
    process.env.agent_id ??
    process.env.connections__service_connection__settings__clientId;

  if (!agentIdentityId || !agentUserOid || !agentBlueprintId) return undefined;

  return { agentIdentityId, agentUserOid, agentBlueprintId };
}
