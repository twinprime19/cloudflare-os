// Sign-in profile seeding: the display name and avatar a gatekeeper offers when a user signs in.
//
// Two levels are covered here. UserDurableObject.loginOrCreateViaGatekeeper() owns the name policy
// (seed on create, backfill only while the stored name is still the local-part seed), and
// LoginConnectCallbackImpl.complete() owns the "hints can never break sign-in" contract: every
// fetch and write is bounded and swallowed, and the session token is still delivered.

import { afterEach, describe, expect, it, vi } from "vitest";
import { env } from "cloudflare:workers";
import { runInDurableObject } from "cloudflare:test";
import type { GatekeeperUser } from "@gadgets/workshop-shared/gatekeeper";
import type { OverseerDurableObject } from "../src/overseer.js";
import { UserDurableObject } from "../src/user.js";
import { LoginConnectCallbackImpl } from "../src/auth/login-flow.js";

declare module "cloudflare:workers" {
  interface ProvidedEnv {
    TEST_OVERSEER: DurableObjectNamespace<OverseerDurableObject>;
  }
}

const EMAIL = "alice.smith@example.com";
const LOCAL_PART = "alice.smith";

const PNG = new Uint8Array([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
const GIF = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]);

type ProfileHints = {
  name?: string;
  photo?: { data: Uint8Array; mimeType: "image/jpeg" | "image/png" };
};

// --- UserDurableObject.loginOrCreateViaGatekeeper ---------------------------------------------

/**
 * Runs `body` against a real UserDurableObject backed by real (empty) Durable Object storage.
 *
 * The runtime only hands out a DurableObjectState from a configured namespace, and the one this
 * package's test worker declares is the Overseer's — so each case borrows a fresh Overseer state and
 * runs a UserDurableObject on it. The state is nothing but storage here; no Overseer method is
 * called. Running inside runInDurableObject() is what makes the DO's synchronous storage legal.
 */
function withUser(body: (user: UserDurableObject) => Promise<void>): Promise<void> {
  const stub = env.TEST_OVERSEER.get(env.TEST_OVERSEER.newUniqueId());
  return runInDurableObject(stub,
      async (_instance: OverseerDurableObject, state: DurableObjectState) => {
        await body(new UserDurableObject(state, env));
      });
}

/** The stored display name, or undefined when the account was never created. */
async function displayName(user: UserDurableObject): Promise<string | undefined> {
  return (await user.whoamiIfExists())?.name;
}

/** Signs a user in with no hint, leaving the display name at the email local-part we seed. */
async function seedLocalPartName(user: UserDurableObject): Promise<void> {
  await user.loginOrCreateViaGatekeeper(EMAIL, true);
}

describe("loginOrCreateViaGatekeeper name seeding", () => {
  it("seeds a new user's display name from the gatekeeper's hint", () => withUser(async user => {
    await expect(user.loginOrCreateViaGatekeeper(EMAIL, true, "Alice Smith")).resolves.toEqual(
        { secret: expect.any(String), nameSeeded: true });

    await expect(displayName(user)).resolves.toBe("Alice Smith");
  }));

  it("trims the hint before seeding", () => withUser(async user => {
    await user.loginOrCreateViaGatekeeper(EMAIL, true, "  Alice Smith  ");

    await expect(displayName(user)).resolves.toBe("Alice Smith");
  }));

  it("falls back to the email local-part when no hint is offered", () => withUser(async user => {
    await expect(user.loginOrCreateViaGatekeeper(EMAIL, true)).resolves.toEqual(
        { secret: expect.any(String), nameSeeded: false });

    await expect(displayName(user)).resolves.toBe(LOCAL_PART);
  }));

  // A gatekeeper is another worker: what it puts in `name` is only a string by declaration.
  it.each([
    ["null", null],
    ["a number", 123],
    ["an object", { toString: () => "Alice Smith" }],
  ])("falls back to the email local-part when the hint is %s", (_label, hint) =>
      withUser(async user => {
    await expect(user.loginOrCreateViaGatekeeper(EMAIL, true, hint)).resolves.toEqual(
        { secret: expect.any(String), nameSeeded: false });

    await expect(displayName(user)).resolves.toBe(LOCAL_PART);
  }));

  it.each([
    ["blank", ""],
    ["whitespace only", "   \t "],
    ["over-long", "A".repeat(101)],
    ["control characters", "Alice\nSmith"],
    ["a null byte", "Alice\u0000Smith"],
  ])("falls back to the email local-part for a %s hint", (_label, hint) => withUser(async user => {
    await user.loginOrCreateViaGatekeeper(EMAIL, true, hint);

    await expect(displayName(user)).resolves.toBe(LOCAL_PART);
  }));

  it("accepts a hint at the length limit", () => withUser(async user => {
    const hint = "A".repeat(100);

    await user.loginOrCreateViaGatekeeper(EMAIL, true, hint);

    await expect(displayName(user)).resolves.toBe(hint);
  }));

  it("does not create an account when signups are closed", () => withUser(async user => {
    await expect(user.loginOrCreateViaGatekeeper(EMAIL, false, "Alice Smith")).resolves.toBeNull();

    await expect(displayName(user)).resolves.toBeUndefined();
  }));

  it("backfills an existing user whose name is still the local-part seed",
      () => withUser(async user => {
    await seedLocalPartName(user);

    await expect(user.loginOrCreateViaGatekeeper(EMAIL, true, "Alice Smith")).resolves.toEqual(
        { secret: expect.any(String), nameSeeded: true });

    await expect(displayName(user)).resolves.toBe("Alice Smith");
  }));

  it("leaves a customised name alone, and says so", () => withUser(async user => {
    await seedLocalPartName(user);
    await user.setOwnDisplayName("Ali");

    await expect(user.loginOrCreateViaGatekeeper(EMAIL, true, "Alice Smith")).resolves.toEqual(
        { secret: expect.any(String), nameSeeded: false });

    await expect(displayName(user)).resolves.toBe("Ali");
  }));

  it("leaves the seeded name alone when the hint is unusable", () => withUser(async user => {
    await seedLocalPartName(user);

    await user.loginOrCreateViaGatekeeper(EMAIL, true, "  ");

    await expect(displayName(user)).resolves.toBe(LOCAL_PART);
  }));

  it("leaves an existing name alone when no hint is offered", () => withUser(async user => {
    await seedLocalPartName(user);
    await user.setOwnDisplayName("Ali");

    await user.loginOrCreateViaGatekeeper(EMAIL, true);

    await expect(displayName(user)).resolves.toBe("Ali");
  }));
});

// --- LoginConnectCallbackImpl.complete ------------------------------------------------------

type LoginCall = { email: string; allowCreate: boolean; nameHint?: string };

/**
 * A LoginConnectCallbackImpl wired to in-memory stand-ins for everything it reaches: the
 * PendingLogin and UserDurableObject bindings, the BLUEPRINTS key holding the admin config, and the
 * AVATARS KV namespace.
 */
function makeCallback(options: {
  providesProfile?: boolean;
  getAuthenticatedProfile?: () => Promise<ProfileHints>;
  existingAvatar?: Uint8Array;
  // What the user DO reports back: false stands for a hint it declined (a customised name).
  nameSeeded?: boolean;
} = {}) {
  const avatars = new Map<string, Uint8Array>();
  if (options.existingAvatar) avatars.set(EMAIL, options.existingAvatar);

  const delivered: string[] = [];
  const failures: string[] = [];
  const loginCalls: LoginCall[] = [];
  let profileCalls = 0;

  const pendingStub = {
    deliver: async (token: string) => { delivered.push(token); },
    fail: async (reason: string) => { failures.push(reason); },
  };
  const userStub = {
    loginOrCreateViaGatekeeper: async (email: string, allowCreate: boolean, nameHint?: string) => {
      loginCalls.push({ email, allowCreate, nameHint });
      return {
        secret: "session-secret",
        nameSeeded: (options.nameSeeded ?? true) && nameHint !== undefined,
      };
    },
  };
  const ctx = {
    props: {
      pendingId: "pending", vendorId: "test", providesProfile: options.providesProfile ?? true,
    },
    exports: {
      PendingLogin: { idFromString: (id: string) => id, get: () => pendingStub },
      UserDurableObject: { idFromName: (name: string) => name, get: () => userStub },
    },
  };
  // BLUEPRINTS holds the admin config (absent = defaults, so signups are open); AVATARS is the KV
  // namespace the seeded photo lands in.
  const callbackEnv = {
    BLUEPRINTS: { get: async () => null },
    AVATARS: {
      put: async (key: string, value: Uint8Array) => { avatars.set(key, value); },
      get: async (key: string, _type: "arrayBuffer") => avatars.get(key)?.buffer ?? null,
    },
  };
  // The counting wrapper hands back the caller's promise as-is rather than awaiting it: an extra
  // async layer would delay handler attachment by a microtask and make a rejection look unhandled.
  const account = {
    getAuthenticatedEmail: async () => EMAIL,
    getAuthenticatedProfile: options.getAuthenticatedProfile === undefined ? undefined : () => {
      profileCalls++;
      return options.getAuthenticatedProfile!();
    },
  };

  type CallbackArgs = ConstructorParameters<typeof LoginConnectCallbackImpl>;
  const callback = new LoginConnectCallbackImpl(
      ctx as unknown as CallbackArgs[0], callbackEnv as unknown as CallbackArgs[1]);

  return {
    callback,
    account: account as unknown as Fetcher<GatekeeperUser>,
    avatar: () => avatars.get(EMAIL),
    delivered, failures, loginCalls,
    profileCalls: () => profileCalls,
  };
}

/** The fields of the one `gatekeeper.login.finished` record emitted during a test. */
function finishedLog(info: { mock: { calls: unknown[][] } }): Record<string, unknown> {
  const records = info.mock.calls
      .map(call => call[0] as Record<string, unknown>)
      .filter(record => record?.event === "gatekeeper.login.finished");
  expect(records).toHaveLength(1);
  return records[0];
}

describe("LoginConnectCallbackImpl profile seeding", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.useRealTimers();
  });

  it("seeds the name and the avatar, then delivers the session", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    const harness = makeCallback({
      getAuthenticatedProfile: async () => ({
        name: "Alice Smith", photo: { data: PNG, mimeType: "image/png" },
      }),
    });

    await harness.callback.complete(harness.account);

    expect(harness.loginCalls).toEqual([
      { email: EMAIL, allowCreate: true, nameHint: "Alice Smith" },
    ]);
    expect(harness.avatar()).toEqual(PNG);
    expect(harness.delivered).toEqual([`${EMAIL}:session-secret`]);
    expect(finishedLog(info)).toMatchObject({ outcome: "ok", profileSeeded: "name+photo" });
  });

  it("keeps the name and the email out of the log line", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    const harness = makeCallback({
      getAuthenticatedProfile: async () => ({ name: "Alice Smith" }),
    });

    await harness.callback.complete(harness.account);

    const logged = JSON.stringify(finishedLog(info));
    expect(logged).not.toContain("Alice Smith");
    expect(logged).not.toContain(EMAIL);
  });

  it("seeds the avatar alone when the gatekeeper offers no name", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    const harness = makeCallback({
      getAuthenticatedProfile: async () => ({ photo: { data: PNG, mimeType: "image/png" } }),
    });

    await harness.callback.complete(harness.account);

    expect(harness.loginCalls).toEqual([{ email: EMAIL, allowCreate: true, nameHint: undefined }]);
    expect(harness.avatar()).toEqual(PNG);
    expect(finishedLog(info)).toMatchObject({ profileSeeded: "photo" });
  });

  it("never overwrites an avatar the user already has", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    const existing = new Uint8Array([0xFF, 0xD8, 0xFF, 0xE0]);
    const harness = makeCallback({
      existingAvatar: existing,
      getAuthenticatedProfile: async () => ({
        name: "Alice Smith", photo: { data: PNG, mimeType: "image/png" },
      }),
    });

    await harness.callback.complete(harness.account);

    expect(harness.avatar()).toEqual(existing);
    expect(harness.delivered).toEqual([`${EMAIL}:session-secret`]);
    expect(finishedLog(info)).toMatchObject({ profileSeeded: "name" });
  });

  it("drops a photo the Workshop won't accept without failing the sign-in", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    vi.spyOn(console, "warn").mockImplementation(() => {});
    const harness = makeCallback({
      getAuthenticatedProfile: async () => ({
        // A GIF: the gatekeeper mislabelled it, so only the magic-byte check catches it.
        name: "Alice Smith", photo: { data: GIF, mimeType: "image/png" },
      }),
    });

    await harness.callback.complete(harness.account);

    expect(harness.avatar()).toBeUndefined();
    expect(harness.delivered).toEqual([`${EMAIL}:session-secret`]);
    expect(finishedLog(info)).toMatchObject({ profileSeeded: "name" });
  });

  it("reports no name seeded when the user DO declines the hint", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    const harness = makeCallback({
      // The stored name was customised, so the DO keeps it and says the hint was not applied.
      nameSeeded: false,
      getAuthenticatedProfile: async () => ({
        name: "Alice Smith", photo: { data: PNG, mimeType: "image/png" },
      }),
    });

    await harness.callback.complete(harness.account);

    expect(harness.loginCalls).toEqual([
      { email: EMAIL, allowCreate: true, nameHint: "Alice Smith" },
    ]);
    expect(harness.delivered).toEqual([`${EMAIL}:session-secret`]);
    expect(finishedLog(info)).toMatchObject({ profileSeeded: "photo" });
  });

  // Everything a gatekeeper returns is another worker's data: a wrong shape must cost the hints,
  // never the sign-in.
  it.each([
    ["a null name", { name: null }],
    ["a numeric name", { name: 123 }],
    ["a null photo", { photo: null }],
    ["a photo that isn't bytes", { photo: { data: "not-bytes", mimeType: "image/png" } }],
    ["no object at all", null],
    ["a string", "Alice Smith"],
  ])("signs the user in when the gatekeeper returns %s", async (_label, hints) => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    vi.spyOn(console, "warn").mockImplementation(() => {});
    const harness = makeCallback({
      getAuthenticatedProfile: async () => hints as unknown as ProfileHints,
    });

    await harness.callback.complete(harness.account);

    expect(harness.loginCalls).toEqual([{ email: EMAIL, allowCreate: true, nameHint: undefined }]);
    expect(harness.avatar()).toBeUndefined();
    expect(harness.delivered).toEqual([`${EMAIL}:session-secret`]);
    expect(harness.failures).toEqual([]);
    expect(finishedLog(info)).toMatchObject({ outcome: "ok", profileSeeded: "none" });
  });

  it("ignores a hint the name policy rejects", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    const harness = makeCallback({
      getAuthenticatedProfile: async () => ({ name: "Alice\nSmith" }),
    });

    await harness.callback.complete(harness.account);

    expect(harness.loginCalls).toEqual([{ email: EMAIL, allowCreate: true, nameHint: undefined }]);
    expect(finishedLog(info)).toMatchObject({ profileSeeded: "none" });
  });

  it("signs the user in when the gatekeeper's profile call throws", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const harness = makeCallback({
      getAuthenticatedProfile: async () => { throw new Error("Graph is down"); },
    });

    await harness.callback.complete(harness.account);

    expect(harness.loginCalls).toEqual([{ email: EMAIL, allowCreate: true, nameHint: undefined }]);
    expect(harness.delivered).toEqual([`${EMAIL}:session-secret`]);
    expect(harness.failures).toEqual([]);
    expect(finishedLog(info)).toMatchObject({ outcome: "ok", profileSeeded: "none" });
    expect(warn).toHaveBeenCalled();
  });

  it("signs the user in when the gatekeeper has no profile method at all", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    vi.spyOn(console, "warn").mockImplementation(() => {});
    const harness = makeCallback();

    await harness.callback.complete(harness.account);

    expect(harness.delivered).toEqual([`${EMAIL}:session-secret`]);
    expect(finishedLog(info)).toMatchObject({ outcome: "ok", profileSeeded: "none" });
  });

  it("does not ask a gatekeeper that doesn't advertise profile hints", async () => {
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    const harness = makeCallback({
      providesProfile: false,
      getAuthenticatedProfile: async () => ({ name: "Alice Smith" }),
    });

    await harness.callback.complete(harness.account);

    expect(harness.profileCalls()).toBe(0);
    expect(harness.loginCalls).toEqual([{ email: EMAIL, allowCreate: true, nameHint: undefined }]);
    expect(finishedLog(info)).toMatchObject({ profileSeeded: "none" });
  });

  it("gives up on the hints rather than the sign-in when the gatekeeper stalls", async () => {
    vi.useFakeTimers();
    const info = vi.spyOn(console, "info").mockImplementation(() => {});
    vi.spyOn(console, "warn").mockImplementation(() => {});
    const harness = makeCallback({
      // Never settles: the gatekeeper accepted the call and then went silent.
      getAuthenticatedProfile: () => new Promise<ProfileHints>(() => {}),
    });

    const done = harness.callback.complete(harness.account);
    await vi.advanceTimersByTimeAsync(5000);
    await done;

    expect(harness.loginCalls).toEqual([{ email: EMAIL, allowCreate: true, nameHint: undefined }]);
    expect(harness.delivered).toEqual([`${EMAIL}:session-secret`]);
    expect(finishedLog(info)).toMatchObject({ outcome: "ok", profileSeeded: "none" });
  });
});
