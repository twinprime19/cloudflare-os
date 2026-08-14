/**
 * The Outlook mailbox is a singleton resource: a connected account has exactly one, so the
 * configurator confirms the choice rather than collecting any values.
 */
export type OutlookMailConfiguratorValues = Record<string, never>;

export interface OutlookMailConfiguratorRpc {}
