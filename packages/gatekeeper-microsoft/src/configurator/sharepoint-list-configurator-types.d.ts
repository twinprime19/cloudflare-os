/**
 * A SharePoint list is named by the address of its page, which the user pastes: the site and list
 * ids Graph works with are not in the browser URL, so the gatekeeper resolves them when the binding
 * is created. The frame therefore collects one value and needs nothing from the gatekeeper.
 */
export type SharePointListConfiguratorValues = {
  /** The pasted list page address. Null while the field is empty: TextInput clears to null. */
  url?: string | null;
}

export interface SharePointListConfiguratorRpc {}
