import { Field, Section, h, type ConfiguratorUISpec } from "@gadgets/configurator-ui";
import type { TeamsConfiguratorRpc, TeamsConfiguratorValues } from "./teams-configurator-types";

// Canonical Teams URL. A connected account has exactly one Teams surface, so this is constant — the
// gatekeeper resolves it against the account behind the connection, not against anything the iframe
// chooses.
const TEAMS_URL = "https://teams.microsoft.com/";

export default {
  initial: {},

  // Nothing to choose, so the connect button is live as soon as the frame opens.
  isReady() {
    return true;
  },

  resourceUrl() {
    return TEAMS_URL;
  },

  render() {
    return <Section>
      <Field
        label="Microsoft Teams"
        description={
          "Connects the Microsoft Teams surface of the account you signed in with. The gadget can " +
          "read the teams and channels you belong to, your chats and their members, the messages " +
          "in both, and search across them. It is read-only: it can never post, edit, delete, or " +
          "join anything."
        }
      />
    </Section>;
  },
} satisfies ConfiguratorUISpec<TeamsConfiguratorRpc, TeamsConfiguratorValues>;
