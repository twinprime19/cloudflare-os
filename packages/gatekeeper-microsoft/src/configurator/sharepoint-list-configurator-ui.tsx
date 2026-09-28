import { Field, h, Section, TextInput, type ConfiguratorUISpec } from "@gadgets/configurator-ui";
import type {
  SharePointListConfiguratorRpc, SharePointListConfiguratorValues,
} from "./sharepoint-list-configurator-types";

/** SharePoint Online's host suffix, matched on a label boundary like the gatekeeper's parser. */
const SHAREPOINT_HOST_SUFFIX = ".sharepoint.com";

/**
 * Is this plausibly the address of a SharePoint list?
 *
 * Only enough to decide whether Connect is live. The same address is parsed strictly, and then
 * resolved against Graph, by the gatekeeper when the binding is created; a check that tried to be
 * authoritative here would either duplicate that parser inside a sandboxed frame or refuse
 * addresses the gatekeeper accepts.
 */
function isPlausibleListUrl(value: string | null | undefined): boolean {
  if (typeof value !== "string") return false;
  let parsed: URL;
  try {
    parsed = new URL(value.trim());
  } catch {
    return false;
  }
  let hostname = parsed.hostname.toLowerCase();
  return parsed.protocol === "https:"
      && hostname.endsWith(SHAREPOINT_HOST_SUFFIX)
      && hostname.length > SHAREPOINT_HOST_SUFFIX.length
      // A list page's address names the list after `/Lists/`. SharePoint URLs are case-preserving
      // but not case-sensitive, so the segment is matched without regard to case.
      && parsed.pathname.toLowerCase().includes("/lists/");
}

export default {
  initial: {},

  isReady({ values }) {
    return isPlausibleListUrl(values.url);
  },

  resourceUrl({ values }) {
    return (values.url ?? "").trim();
  },

  render({ values, setValues }) {
    return <Section>
      <Field
        label="List URL"
        description={
          "Open the list in SharePoint and paste its address, e.g. " +
          "https://contoso.sharepoint.com/sites/HR/Lists/Requests/AllItems.aspx"
        }
      >
        <TextInput
          name="url"
          value={values.url}
          placeholder="https://contoso.sharepoint.com/sites/HR/Lists/Requests/AllItems.aspx"
          onChange={url => setValues({ url })}
        />
      </Field>
    </Section>;
  },
} satisfies ConfiguratorUISpec<SharePointListConfiguratorRpc, SharePointListConfiguratorValues>;
