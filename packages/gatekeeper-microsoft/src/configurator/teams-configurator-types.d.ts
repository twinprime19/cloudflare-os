/**
 * Microsoft Teams is a singleton resource: a connected account has exactly one Teams surface, so
 * the configurator confirms the choice rather than collecting any values.
 */
export type TeamsConfiguratorValues = Record<string, never>;

export interface TeamsConfiguratorRpc {}
