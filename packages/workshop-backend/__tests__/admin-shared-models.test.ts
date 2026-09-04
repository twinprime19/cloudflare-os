// The deployment-wide model catalog an admin curates, and what users see of it.
//
// Two implementations meet here, both run for real:
//   * AdminSettings owns the catalog -- addSharedModel/deleteSharedModel/setDefaultModel/
//     setModelEnabled -- writing the shared-model rows into its own storage and mirroring the
//     admin config to KV, which every hot path reads instead of waking the singleton.
//   * UserDurableObject.listModels() and getChatContext() consume it: the catalog minus the
//     models an admin turned off, ordered with the deployment default first, and a per-turn
//     resolution that refuses a disabled model and looks a shared one up by id.
//
// admin-config.test.ts covers the parse/serialize half of the KV mirror; the RPC implementations
// that write it, and the two user-facing readers, are what these pin.
//
// Both durable objects are constructed on one borrowed Durable Object state, so their real
// storage is real -- only the DO *namespace* the user DO reaches its settings singleton through
// is stood in, pointing at the AdminSettings instance in the same isolate. KV is a Map, since the
// unit test worker binds no namespace (the same approach login-profile.test.ts takes).

import { afterEach, describe, expect, it, vi } from "vitest";
import { env } from "cloudflare:workers";
import { runInDurableObject } from "cloudflare:test";
import type { AiChatAuthorInfo, AiModelConfig } from "@gadgets/workshop-shared/api";
import { SUGGESTED_MODELS } from "@gadgets/workshop-shared/api";
import type { AdminConfig } from "../src/admin-config.js";
import { parseAdminConfig } from "../src/admin-config.js";
import { ADMIN_CONFIG_KEY } from "../src/blueprint-archive.js";
import { AdminSettings } from "../src/admin-settings.js";
import { UserDurableObject } from "../src/user.js";
import type { OverseerDurableObject } from "../src/overseer.js";

declare module "cloudflare:workers" {
  interface ProvidedEnv {
    TEST_OVERSEER: DurableObjectNamespace<OverseerDurableObject>;
  }
}

const SHARED_PROFILE: AiChatAuthorInfo = { type: "agent", id: "acme/model-1", name: "Acme One" };
const SHARED_CONFIG: AiModelConfig = {
  provider: "anthropic", model: "acme-1", apiToken: "acme-secret-token",
};

// The deployment's default suggested Cloudflare model -- first in SUGGESTED_MODELS.cloudflare,
// so a new chat starts on it when the admin has not set a default. Its id is the deployment's
// own decision, so it stays literal: renaming a model upstream must fail this test.
const BUILT_IN_ID = "@cf/zai-org/glm-5.3-flash";
const BUILT_IN_NAME = "GLM 5.3 Flash (Workers AI)";

// A built-in from a provider whose catalog is not ours to pin: read out of the catalog rather
// than named, so renaming a model upstream cannot fail a test about fork behaviour.
const ANTHROPIC_BUILT_IN = Object.keys(SUGGESTED_MODELS.anthropic)[0];

afterEach(() => {
  vi.restoreAllMocks();
});

type Harness = {
  admin: AdminSettings;
  user: UserDurableObject;
  /** The admin config as the KV mirror holds it -- what every hot path actually reads. */
  mirroredConfig: () => AdminConfig;
  /** Point the user DO at a settings singleton that is down. */
  breakSettings: () => void;
};

type SettingsNamespaceHolder = { adminSettings: { getByName: (name: string) => unknown } };

let doCounter = 0;

/**
 * Runs `body` against a real AdminSettings and a real UserDurableObject wired to each other.
 *
 * `providers` enables AI Gateway mode with that provider list; omit it for a deployment with no
 * gateway, where the catalog is the shared models alone.
 */
function withDeployment(
    options: { providers?: string },
    body: (harness: Harness) => Promise<void>): Promise<void> {
  const stub = env.TEST_OVERSEER.getByName(`admin-shared-models-${++doCounter}`);
  return runInDurableObject(stub,
      async (_instance: OverseerDurableObject, state: DurableObjectState) => {
        const kv = new Map<string, string>();
        const deploymentEnv = {
          ...(options.providers === undefined ? {} : {
            CF_AI_GATEWAY: "platform-gateway",
            CF_AI_GATEWAY_ACCOUNT_ID: "gateway-account-id",
            CF_AI_GATEWAY_API_TOKEN: "gateway-token",
            CF_AI_GATEWAY_PROVIDERS: options.providers,
          }),
          BLUEPRINTS: {
            get: async (key: string) => kv.get(key) ?? null,
            put: async (key: string, value: string) => { kv.set(key, value); },
          },
        } as unknown as Cloudflare.Env;

        const admin = new AdminSettings(state, deploymentEnv);
        const user = new UserDurableObject(state, deploymentEnv);
        const holder = user as unknown as SettingsNamespaceHolder;
        holder.adminSettings = { getByName: () => admin };

        await body({
          admin,
          user,
          mirroredConfig: () => parseAdminConfig(kv.get(ADMIN_CONFIG_KEY) ?? null),
          breakSettings: () => {
            holder.adminSettings = { getByName: () => ({
              listSharedModels: async () => { throw new Error("settings DO unavailable"); },
              getSharedModelRecord: async () => { throw new Error("settings DO unavailable"); },
            }) };
          },
        });
      });
}

/** Adds the deployment's one shared model. */
async function addSharedModel(admin: AdminSettings): Promise<void> {
  await admin.addSharedModel({ ...SHARED_PROFILE }, { ...SHARED_CONFIG });
}

/**
 * Writes a shared-model row directly into the catalog, for states addSharedModel refuses to
 * create today but that a deployment may already hold.
 */
function seedSharedModel(
    admin: AdminSettings, profile: AiChatAuthorInfo, config: AiModelConfig): void {
  (admin as unknown as { storage: { sharedAiModels: { put: (record: unknown) => void } } })
      .storage.sharedAiModels.put({ profile, config });
}

async function listedIds(user: UserDurableObject): Promise<string[]> {
  return (await user.listModels()).map(entry => entry.id);
}

describe("AdminSettings shared-model catalog", () => {
  it("offers an added model to every user, keeping its configuration out of the listing",
      () => withDeployment({}, async ({ admin }) => {
    await addSharedModel(admin);

    expect(admin.listSharedModels()).toEqual([{ ...SHARED_PROFILE, type: "agent" }]);
    // The configuration -- API token included -- is readable only through the backend-only
    // accessor, never from the published profiles.
    expect(admin.getSharedModelRecord(SHARED_PROFILE.id)).toEqual({
      profile: { ...SHARED_PROFILE, type: "agent" },
      config: SHARED_CONFIG,
    });
  }));

  it("replaces a shared model added again under the same id",
      () => withDeployment({}, async ({ admin }) => {
    await addSharedModel(admin);
    await admin.addSharedModel(
        { type: "agent", id: SHARED_PROFILE.id, name: "Acme One (rotated)" },
        { ...SHARED_CONFIG, apiToken: "rotated-token" });

    expect(admin.listSharedModels()).toHaveLength(1);
    expect(admin.listSharedModels()[0].name).toBe("Acme One (rotated)");
    expect(admin.getSharedModelRecord(SHARED_PROFILE.id)?.config.apiToken).toBe("rotated-token");
  }));

  it("stores a model classed as an agent whatever the caller declared",
      () => withDeployment({}, async ({ admin }) => {
    await admin.addSharedModel(
        { type: "user", id: SHARED_PROFILE.id, name: "Acme One" } as AiChatAuthorInfo,
        { ...SHARED_CONFIG });

    expect(admin.listSharedModels()[0].type).toBe("agent");
  }));

  it("records the deployment default in the durable object and its KV mirror alike",
      () => withDeployment({}, async ({ admin, mirroredConfig }) => {
    await addSharedModel(admin);

    await admin.setDefaultModel(SHARED_PROFILE.id);

    expect(admin.getAdminConfig().defaultModelId).toBe(SHARED_PROFILE.id);
    expect(mirroredConfig().defaultModelId).toBe(SHARED_PROFILE.id);

    await admin.setDefaultModel(null);

    expect(mirroredConfig().defaultModelId).toBeNull();
  }));

  it("refuses a default naming a model no user can reach",
      () => withDeployment({}, async ({ admin, mirroredConfig }) => {
    await expect(admin.setDefaultModel("acme/typo")).rejects.toThrow("No such model: acme/typo");

    expect(mirroredConfig().defaultModelId).toBeNull();
  }));

  it("turns a shared model off and on again through the KV mirror",
      () => withDeployment({}, async ({ admin, mirroredConfig }) => {
    await addSharedModel(admin);

    await admin.setModelEnabled(SHARED_PROFILE.id, false);

    expect(admin.getAdminConfig().disabledModelIds).toEqual([SHARED_PROFILE.id]);
    expect(mirroredConfig().disabledModelIds).toEqual([SHARED_PROFILE.id]);

    await admin.setModelEnabled(SHARED_PROFILE.id, true);

    expect(mirroredConfig().disabledModelIds).toEqual([]);
  }));

  // Checked in both directions, so an enable cannot quietly file away a typo either.
  it("refuses to turn an unknown model off or on",
      () => withDeployment({}, async ({ admin, mirroredConfig }) => {
    await expect(admin.setModelEnabled("acme/typo", false))
        .rejects.toThrow("No such model: acme/typo");
    await expect(admin.setModelEnabled("acme/typo", true))
        .rejects.toThrow("No such model: acme/typo");

    expect(mirroredConfig().disabledModelIds).toEqual([]);
  }));

  // Deletion is permanent, so a leftover denylist entry would silently resurrect the disabled
  // state if an admin later re-added a model under the same id.
  it("clears the default and the denylist when a shared model is deleted",
      () => withDeployment({}, async ({ admin, mirroredConfig }) => {
    await addSharedModel(admin);
    await admin.setDefaultModel(SHARED_PROFILE.id);
    await admin.setModelEnabled(SHARED_PROFILE.id, false);

    await admin.deleteSharedModel(SHARED_PROFILE.id);

    expect(admin.listSharedModels()).toEqual([]);
    expect(admin.getSharedModelRecord(SHARED_PROFILE.id)).toBeUndefined();
    expect(mirroredConfig().defaultModelId).toBeNull();
    expect(mirroredConfig().disabledModelIds).toEqual([]);
  }));

  it("leaves another model's default alone when a shared model is deleted",
      () => withDeployment({ providers: "cloudflare,anthropic" },
          async ({ admin, mirroredConfig }) => {
    await addSharedModel(admin);
    await admin.setDefaultModel(BUILT_IN_ID);

    await admin.deleteSharedModel(SHARED_PROFILE.id);

    expect(mirroredConfig().defaultModelId).toBe(BUILT_IN_ID);
  }));

  it("refuses a shared model whose provider the gateway does not serve",
      () => withDeployment({ providers: "cloudflare" }, async ({ admin }) => {
    await expect(admin.addSharedModel({ ...SHARED_PROFILE }, { ...SHARED_CONFIG }))
        .rejects.toThrow('Provider "anthropic" is not available in AI Gateway mode.');

    expect(admin.listSharedModels()).toEqual([]);
  }));

  // A built-in's id resolves from the gateway catalog first, so a shared entry reusing one would
  // be listed under the built-in's name and never actually be used.
  it("refuses a shared model that reuses a built-in's id",
      () => withDeployment({ providers: "cloudflare" }, async ({ admin }) => {
    await expect(admin.addSharedModel(
        { type: "agent", id: BUILT_IN_ID, name: "Impostor" },
        { provider: "cloudflare", model: BUILT_IN_ID, apiToken: "token" },
    )).rejects.toThrow(`Model id "${BUILT_IN_ID}" belongs to built-in model "${BUILT_IN_NAME}".`);
  }));

  it("turns a gateway built-in off deployment-wide",
      () => withDeployment({ providers: "cloudflare" }, async ({ admin, mirroredConfig }) => {
    await admin.setModelEnabled(BUILT_IN_ID, false);

    expect(mirroredConfig().disabledModelIds).toEqual([BUILT_IN_ID]);
  }));
});

describe("UserDurableObject.listModels", () => {
  it("drops the models an admin turned off, built-ins included",
      () => withDeployment({ providers: "cloudflare,anthropic" }, async ({ admin, user }) => {
    await addSharedModel(admin);
    await admin.setModelEnabled(ANTHROPIC_BUILT_IN, false);
    await admin.setModelEnabled(SHARED_PROFILE.id, false);

    const ids = await listedIds(user);

    expect(ids).not.toContain(ANTHROPIC_BUILT_IN);
    expect(ids).not.toContain(SHARED_PROFILE.id);
    expect(ids).toContain(BUILT_IN_ID);
  }));

  it("lists the deployment default first, because a new chat starts on the first entry",
      () => withDeployment({ providers: "cloudflare,anthropic" }, async ({ admin, user }) => {
    await addSharedModel(admin);
    const natural = await listedIds(user);
    expect(natural[0]).not.toBe(SHARED_PROFILE.id);

    await admin.setDefaultModel(SHARED_PROFILE.id);

    const ordered = await listedIds(user);
    expect(ordered[0]).toBe(SHARED_PROFILE.id);
    expect(ordered.slice(1).toSorted()).toEqual(
        natural.filter(id => id !== SHARED_PROFILE.id).toSorted());
  }));

  it("moves a built-in to the front when it is the deployment default",
      () => withDeployment({ providers: "cloudflare,anthropic" }, async ({ admin, user }) => {
    const natural = await listedIds(user);
    expect(natural[0]).not.toBe(ANTHROPIC_BUILT_IN);

    await admin.setDefaultModel(ANTHROPIC_BUILT_IN);

    expect((await listedIds(user))[0]).toBe(ANTHROPIC_BUILT_IN);
  }));

  it("keeps the natural order while the default model is turned off",
      () => withDeployment({ providers: "cloudflare,anthropic" }, async ({ admin, user }) => {
    const natural = await listedIds(user);
    await admin.setDefaultModel(ANTHROPIC_BUILT_IN);
    await admin.setModelEnabled(ANTHROPIC_BUILT_IN, false);

    expect(await listedIds(user))
        .toEqual(natural.filter(id => id !== ANTHROPIC_BUILT_IN));
  }));

  // The gateway's entry is the one getChatContext resolves, so a shared entry reusing that id
  // would name a model nobody can reach.
  it("skips a shared model that duplicates a gateway built-in",
      () => withDeployment({ providers: "cloudflare" }, async ({ admin, user }) => {
    // Written straight into the catalog, because addSharedModel refuses this id while the gateway
    // serves the provider -- which is the point: the row can only be one an older deployment
    // stored before the gateway (or the guard) reached it, and the listing has to hold the line
    // for it anyway.
    seedSharedModel(admin,
        { type: "agent", id: BUILT_IN_ID, name: "Impostor" },
        { provider: "anthropic", model: "impostor", apiToken: "token" });

    const listed = await user.listModels();

    expect(listed.filter(entry => entry.id === BUILT_IN_ID))
        .toEqual([{ type: "agent", id: BUILT_IN_ID, name: BUILT_IN_NAME }]);
  }));

  it("still lists the built-ins when the settings singleton is down",
      () => withDeployment({ providers: "cloudflare" }, async ({ user, breakSettings }) => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    breakSettings();

    expect(await listedIds(user)).toEqual(Object.keys(SUGGESTED_MODELS.cloudflare));
    expect(warn.mock.calls.map(call => (call[0] as { event?: string })?.event))
        .toContain("shared.models.list.failed");
  }));
});

describe("UserDurableObject.getChatContext", () => {
  it("resolves a shared model, configuration and all, for the turn",
      () => withDeployment({}, async ({ admin, user }) => {
    await addSharedModel(admin);

    const context = await user.getChatContext(SHARED_PROFILE.id);

    expect(context.aiModel).toEqual({
      profile: { ...SHARED_PROFILE, type: "agent" },
      config: SHARED_CONFIG,
    });
  }));

  // A chat still pinned to a model the admin turned off fails here rather than quietly carrying
  // on with a model that has left every catalog.
  it("refuses a model an admin turned off", () => withDeployment({}, async ({ admin, user }) => {
    await addSharedModel(admin);
    await admin.setModelEnabled(SHARED_PROFILE.id, false);

    await expect(user.getChatContext(SHARED_PROFILE.id))
        .rejects.toThrow(`Model disabled by administrator: ${SHARED_PROFILE.id}`);
    // The guard is about the denylist alone: a model that was never in the catalog still reads as
    // unknown, rather than being reported as something an admin turned off.
    await expect(user.getChatContext("acme/typo")).rejects.toThrow("No such model: acme/typo");
  }));

  it("refuses a disabled built-in before the gateway ever resolves it",
      () => withDeployment({ providers: "cloudflare" }, async ({ admin, user }) => {
    await admin.setModelEnabled(BUILT_IN_ID, false);

    await expect(user.getChatContext(BUILT_IN_ID))
        .rejects.toThrow(`Model disabled by administrator: ${BUILT_IN_ID}`);
  }));
});
