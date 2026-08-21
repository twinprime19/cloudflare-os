import type { AiChatAuthorInfo } from "@gadgets/workshop-shared/api";

/** Sentinel used for UI values so an explicit "no agent" choice has a select value of its own. */
export const NO_AGENT_OPTION_VALUE = "__gadgets_no_agent__";

/**
 * The model a new chat starts on: the first the server lists, which is the model the deployment's
 * admin made default.
 *
 * Nothing is remembered between chats. A user switching models in the picker changes that chat and
 * only that chat, so the admin's default is what every new chat gets, in every browser.
 */
export function getStoredSelectedModel(
  models: AiChatAuthorInfo[],
): string | null {
  return models[0]?.id ?? null;
}

export function toModelSelectValue(modelId: string | null): string {
  return modelId ?? NO_AGENT_OPTION_VALUE;
}

export function fromModelSelectValue(value: string): string | null {
  return value === NO_AGENT_OPTION_VALUE ? null : value;
}
