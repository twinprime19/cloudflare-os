// The deployment's AI model catalog, end to end.
//
// Models used to be per-user: whatever a user added, only that user could see. They are the
// deployment's now -- an admin curates one catalog and every user chats with it -- and the two
// halves of that live in different durable objects. The admin writes the AdminSettings singleton;
// each user's own object reads it when listing models. Nothing but a real run exercises that hop,
// so the property that matters ("what the admin added, a *different* user sees") is only observable
// here: a unit test for either object alone sees its own side and stubs the other.
//
// Nothing is stubbed but the network. The real workshop-backend runs under wrangler, and the tests
// speak Cap'n Web over a WebSocket to /api exactly as the browser does. No test sends a chat turn,
// so the stand-in API tokens below are never used against a provider.
//
// These cases share the deployment's one catalog, so they are deliberately NOT `it.concurrent`:
// ordering assertions read a setting the whole deployment shares. They stay independent by taking a
// fresh model id per case, the way the rest of the suite takes fresh usernames.

import { afterAll, beforeAll, describe, expect, it } from "vitest";
import type { RpcStub } from "capnweb";
import type {
  AdminApi, AiChatAuthorInfo, AiModelConfig, AuthenticatedApi, PublicApi,
} from "@gadgets/workshop-shared/api";
import { ADMIN_USERNAME, startHarness, type Harness } from "../src/harness.js";
import { connect, nextUsernames, signUp } from "../src/rpc-client.js";
import { NetworkInterceptor } from "../src/network-interceptor.js";

/** Stand-in provider credential. Never used: nothing here asks a model to generate anything. */
const MODEL_TOKEN = "sk-integration-test-not-a-real-key";

let harness: Harness;
let interceptor: NetworkInterceptor;
/** The admin's session, held open for the file: `adminApi` below is a capability minted on it. */
let adminSession: RpcStub<PublicApi>;
let adminApi: RpcStub<AdminApi>;

beforeAll(async () => {
  // No handlers: nothing here should make an outbound request at all, so every one is a failure.
  interceptor = new NetworkInterceptor();
  interceptor.install();
  harness = await startHarness({
    // No gatekeeper: this suite signs up with a password and never connects an account.
    gatekeepers: [],
    // Pin AI Gateway mode off. A developer's root .dev.vars can switch it on for the Workshop the
    // harness boots (docs/integration-testing.md), which would prepend the gateway's built-in
    // models to every listModels() result -- so the ordering assertions below would pass on CI and
    // mean something different on that machine. AUTH_GATEKEEPERS goes with it: an allowlisted
    // vendor from the same file names a gatekeeper this harness does not bind.
    workshopVars: { CF_AI_GATEWAY: "", AUTH_GATEKEEPERS: "" },
  });

  adminSession = connect(harness.url);
  const admin = await signUp(adminSession, ADMIN_USERNAME);
  const api = await admin.getAdminApi();
  if (!api) {
    throw new Error(
      `"${ADMIN_USERNAME}" was not granted admin rights -- vars.ADMINS did not reach the Worker.`);
  }
  adminApi = api;
});

afterAll(async () => {
  const unmocked = interceptor.getUnmockedCalls();
  adminSession?.[Symbol.dispose]();
  await harness?.server.close();
  interceptor.uninstall();
  interceptor.reset();
  expect(unmocked).toEqual([]);
});

/** Sign up a brand-new (non-admin) user on their own session and act as them. */
async function withNewUser<T>(
    prefix: string, body: (api: RpcStub<AuthenticatedApi>) => Promise<T>): Promise<T> {
  const publicApi = connect(harness.url);
  try {
    const [username] = nextUsernames(prefix);
    using api = await signUp(publicApi, username);
    return await body(api);
  } finally {
    publicApi[Symbol.dispose]();
  }
}

let modelSeq = 0;

/** A model no other case in this file can name, so the shared catalog can't collide. */
function freshModel(): { profile: AiChatAuthorInfo; config: AiModelConfig } {
  const n = ++modelSeq;
  return {
    profile: { type: "agent", id: `integration-test/model-${n}`, name: `Test Model ${n}` },
    config: { provider: "openai", model: `test-model-${n}`, apiToken: MODEL_TOKEN },
  };
}

describe("shared AI models", () => {
  it("refuses a non-admin's attempt to add a model", async () => {
    const { profile, config } = freshModel();

    await withNewUser("nonadmin", async api => {
      // No admin capability to reach the shared catalog with...
      await expect(api.getAdminApi()).resolves.toBeNull();
      // ...and the legacy per-user path is closed too, so there is no way in at all.
      await expect(api.addModel(profile, config)).rejects.toThrow(/administrator/i);
      await expect(api.listModels()).resolves.not.toContainEqual(
        expect.objectContaining({ id: profile.id }));
    });
  });

  it("offers a model the admin adds to a different user", async () => {
    const { profile, config } = freshModel();
    await adminApi.addSharedModel(profile, config);

    await withNewUser("viewer", async api => {
      // Compared whole rather than by id: a profile carries exactly these three fields, so an
      // extra one -- the provider API token above, say -- fails here rather than reaching a client.
      await expect(api.listModels()).resolves.toContainEqual(
        { type: "agent", id: profile.id, name: profile.name });
    });
  });

  it("offers a renamed model under its new name to a different user", async () => {
    const { profile, config } = freshModel();
    await adminApi.addSharedModel(profile, config);

    await adminApi.renameSharedModel(profile.id, "Renamed Model");

    await withNewUser("renamed", async api => {
      const models = await api.listModels();
      const renamed = models.find(model => model.id === profile.id);
      // The whole profile again, because renaming re-puts the stored record: were the write to
      // rebuild the entry from the configuration instead of carrying it through untouched, the
      // provider API token would surface here rather than staying on the server.
      expect(renamed).toEqual({ type: "agent", id: profile.id, name: "Renamed Model" });
    });
  });

  it("refuses to rename a model no user could reach anyway", async () => {
    await expect(adminApi.renameSharedModel("integration-test/never-added", "X"))
      .rejects.toThrow(/no such model/i);
  });

  it("refuses a blank new name, and stores a padded one trimmed", async () => {
    const { profile, config } = freshModel();
    await adminApi.addSharedModel(profile, config);

    await expect(adminApi.renameSharedModel(profile.id, "")).rejects.toThrow(/cannot be empty/i);
    await expect(adminApi.renameSharedModel(profile.id, "   ")).rejects.toThrow(/cannot be empty/i);
    // A rejected rename leaves the catalog entry as it was, name included.
    await withNewUser("unrenamed", async api => {
      await expect(api.listModels()).resolves.toContainEqual(
        { type: "agent", id: profile.id, name: profile.name });
    });

    await adminApi.renameSharedModel(profile.id, "  Padded Name  ");

    await withNewUser("padded", async api => {
      await expect(api.listModels()).resolves.toContainEqual(
        { type: "agent", id: profile.id, name: "Padded Name" });
    });
  });

  it("lists the deployment's default model first for every user", async () => {
    const first = freshModel();
    const second = freshModel();
    await adminApi.addSharedModel(first.profile, first.config);
    await adminApi.addSharedModel(second.profile, second.config);

    // The later-added model, so being first can only be the default's doing and not insertion order.
    await adminApi.setDefaultModel(second.profile.id);

    await withNewUser("chatter", async api => {
      const models = await api.listModels();
      expect(models[0]).toEqual(
        { type: "agent", id: second.profile.id, name: second.profile.name });
      expect(models).toContainEqual(
        { type: "agent", id: first.profile.id, name: first.profile.name });
    });
  });

  it("clears the default, and stops offering the model, when the default is deleted", async () => {
    const { profile, config } = freshModel();
    await adminApi.addSharedModel(profile, config);
    await adminApi.setDefaultModel(profile.id);
    await expect(adminApi.getSettings()).resolves.toMatchObject({ defaultModelId: profile.id });

    await adminApi.deleteSharedModel(profile.id);

    // A stale default would leave every user's catalog ordered around a model that is gone.
    await expect(adminApi.getSettings()).resolves.toMatchObject({ defaultModelId: null });
    await withNewUser("latecomer", async api => {
      await expect(api.listModels()).resolves.not.toContainEqual(
        expect.objectContaining({ id: profile.id }));
    });
  });

  it("takes a model the admin turns off out of every user's catalog", async () => {
    const { profile, config } = freshModel();
    await adminApi.addSharedModel(profile, config);

    await adminApi.setModelEnabled(profile.id, false);

    await withNewUser("shutout", async api => {
      await expect(api.listModels()).resolves.not.toContainEqual(
        expect.objectContaining({ id: profile.id }));

      // Turned off, not merely hidden. A workspace binding re-resolves the model when it is
      // instantiated, which is the one place a user can still name an id the list no longer
      // offers -- and it fails there rather than quietly binding a model nobody may use.
      using gadget = await api.newGadget();
      await expect(gadget.newAiModelGatekeeper(profile.id))
        .rejects.toThrow(/disabled by administrator/i);
    });
  });

  it("offers a model again once the admin turns it back on", async () => {
    const { profile, config } = freshModel();
    await adminApi.addSharedModel(profile, config);
    await adminApi.setModelEnabled(profile.id, false);

    await adminApi.setModelEnabled(profile.id, true);

    await withNewUser("returning", async api => {
      await expect(api.listModels()).resolves.toContainEqual(
        { type: "agent", id: profile.id, name: profile.name });
    });
  });

  it("keeps the deployment's default while the default model is turned off", async () => {
    const { profile, config } = freshModel();
    await adminApi.addSharedModel(profile, config);
    await adminApi.setDefaultModel(profile.id);

    await adminApi.setModelEnabled(profile.id, false);

    // Unlike deleting it, turning the default off is temporary, so the setting stays.
    await expect(adminApi.getSettings()).resolves.toMatchObject({ defaultModelId: profile.id });
    await withNewUser("displaced", async api => {
      // A default nobody can see just leaves the catalog in its natural order: no error, no gap.
      const models = await api.listModels();
      expect(models).not.toContainEqual(expect.objectContaining({ id: profile.id }));
      expect(models.length).toBeGreaterThan(0);
    });

    await adminApi.setModelEnabled(profile.id, true);

    // ...and default-first ordering comes back on its own, with no second setDefaultModel call.
    await withNewUser("restored", async api => {
      const models = await api.listModels();
      expect(models[0]).toEqual({ type: "agent", id: profile.id, name: profile.name });
    });
  });

  it("refuses a non-admin's attempt to turn a model back on", async () => {
    const { profile, config } = freshModel();
    await adminApi.addSharedModel(profile, config);
    await adminApi.setModelEnabled(profile.id, false);

    await withNewUser("nonadmin", async api => {
      // setModelEnabled is only reachable through the admin capability, and a non-admin is never
      // handed one -- so a user cannot undo the admin's decision from their own session.
      await expect(api.getAdminApi()).resolves.toBeNull();
      await expect(api.listModels()).resolves.not.toContainEqual(
        expect.objectContaining({ id: profile.id }));
    });
  });

  it("refuses to disable a model no user could reach anyway", async () => {
    await expect(adminApi.setModelEnabled("integration-test/never-added", false))
      .rejects.toThrow(/no such model/i);
  });
});
