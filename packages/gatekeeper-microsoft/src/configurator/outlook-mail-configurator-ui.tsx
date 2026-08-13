import { Field, Section, h, type ConfiguratorUISpec } from "@gadgets/configurator-ui";
import type {
  OutlookMailConfiguratorRpc, OutlookMailConfiguratorValues,
} from "./outlook-mail-configurator-types";

// Canonical mailbox URL. A connected account has exactly one mailbox, so this is constant — the
// gatekeeper resolves it against the account behind the connection, not against anything the
// iframe chooses.
const OUTLOOK_MAIL_URL = "https://outlook.office.com/mail/";

export default {
  initial: {},

  // Nothing to choose, so the connect button is live as soon as the frame opens.
  isReady() {
    return true;
  },

  resourceUrl() {
    return OUTLOOK_MAIL_URL;
  },

  render() {
    return <Section>
      <Field
        label="Outlook mailbox"
        description={
          "Connects the mailbox of the Microsoft account you signed in with. The gadget can read " +
          "messages and folders; marking messages read, moving them, and creating reply drafts " +
          "are queued for your approval. It can never send mail."
        }
      />
    </Section>;
  },
} satisfies ConfiguratorUISpec<OutlookMailConfiguratorRpc, OutlookMailConfiguratorValues>;
