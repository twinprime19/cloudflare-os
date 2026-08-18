// Sign-in through an authentication gatekeeper, end to end.
//
// The Workshop takes three things off a gatekeeper at sign-in: the verified email it keys the user
// by, a display-name hint, and a profile photo. The first has always been covered only by unit tests
// with a stubbed account; the photo is the one place a real Uint8Array crosses the Worker boundary
// from a gatekeeper into core storage, which no unit test can exercise -- a mock hands back the same
// object it was given, while here the bytes are serialised over RPC, validated, written to the AVATARS
// KV namespace, and read back out over the browser's own API.
//
// Nothing is stubbed but the network. The real workshop-backend runs under wrangler, tests speak Cap'n
// Web over a WebSocket to /api exactly as the browser does, and the gatekeeper is a real Worker
// speaking the real protocol: startGatekeeperLogin() gets a URL from its connectAccount(), fetching
// that URL completes the flow, and the session token comes back through the same PendingLogin
// rendezvous the browser waits on. The one thing the fixture stands in for is the provider's login
// page -- `?email=` on the URL is where a human would pick an account.

import { afterAll, beforeAll, describe, expect, it } from "vitest";
import type { RpcStub } from "capnweb";
import type { AuthenticatedApi, PublicApi } from "@gadgets/workshop-shared/api";
import {
  startHarness, TEST_GATEKEEPER_BINDING, TEST_GATEKEEPER_DIR, TEST_GATEKEEPER_WORKER,
  TEST_VENDOR_ID, type Harness,
} from "../src/harness.js";
import { connect, nextUsernames } from "../src/rpc-client.js";
import { NetworkInterceptor } from "../src/network-interceptor.js";

// The fixture's fixed display-name hint (fixtures/gatekeeper-test/src/test-gatekeeper.ts). Mirrored
// here rather than imported: a Worker entry module can export nothing but its entrypoints. The photo
// is read back off the fixture instead, so the bytes are compared against the ones it really sent.
const PROFILE_NAME = "Test Gatekeeper User";

// A PNG header plus filler, for the cases that need an avatar the gatekeeper did not supply. The
// Workshop checks the magic bytes and the size, never the pixels.
const CUSTOM_AVATAR = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

let harness: Harness;
let interceptor: NetworkInterceptor;

beforeAll(async () => {
  // No handlers: signing in through this gatekeeper must not touch the network at all.
  interceptor = new NetworkInterceptor();
  interceptor.install();
  harness = await startHarness({
    gatekeepers: [{ binding: TEST_GATEKEEPER_BINDING, dir: TEST_GATEKEEPER_DIR }],
    // Sign-in is opt-in per deployment: without the allowlist, startGatekeeperLogin() refuses however
    // auth-capable the bound vendor is.
    workshopVars: { AUTH_GATEKEEPERS: TEST_VENDOR_ID },
  });
});

afterAll(async () => {
  // Asserted once for the whole file rather than per test: the tests run concurrently, so an
  // afterEach would inspect and clear state its siblings are still using.
  const unmocked = interceptor.getUnmockedCalls();
  await harness?.server.close();
  interceptor.uninstall();
  interceptor.reset();
  expect(unmocked).toEqual([]);
});

// Each test gets its own RPC session, so a disposal in one can't disturb another running alongside.
async function withSession<T>(body: (api: RpcStub<PublicApi>) => Promise<T>): Promise<T> {
  const publicApi = connect(harness.url);
  try {
    return await body(publicApi);
  } finally {
    publicApi[Symbol.dispose]();
  }
}

/** A sign-in identity no other test can pick. */
function signInEmail(prefix: string): string {
  const [name] = nextUsernames(prefix);
  return `${name}@gadgets-test.example`;
}

/** The hints the fixture gatekeeper hands the Workshop at sign-in, from the fixture itself. */
async function fixtureProfileHints(): Promise<{ name: string; photo: Uint8Array }> {
  const res = await harness.fetchWorker(
    TEST_GATEKEEPER_WORKER, "http://gatekeeper-test.test/control/profile-hints");
  if (res.status !== 200) {
    throw new Error(
      `Reading the fixture's profile hints failed with ${res.status}: ${await res.text()}`);
  }
  const { name, photoBase64 } = await res.json() as { name: string; photoBase64: string };
  return { name, photo: new Uint8Array(Buffer.from(photoBase64, "base64")) };
}

/**
 * Sign in as `email` and return the session token, driving the steps the browser drives: start the
 * attempt, begin waiting on it, then open the gatekeeper's pop-up URL.
 *
 * The wait is registered before the pop-up completes, as it is in the browser -- so the token is
 * delivered to a live waiter rather than stashed for a later arrival.
 */
async function signIn(publicApi: RpcStub<PublicApi>, email: string): Promise<string> {
  const { url, attempt } = await publicApi.startGatekeeperLogin(TEST_VENDOR_ID);
  try {
    const waiting = attempt.wait();
    // If the pop-up step below throws, this call is abandoned mid-flight and rejects when the session
    // is disposed. Mark it handled so that rejection can't surface as a stray unhandled error on top
    // of the real failure; `await waiting` still sees it on the path that gets there.
    waiting.catch(() => {});
    // The fixture's stand-in for the provider's login page: who completes the flow is a query param.
    const popup = new URL(url);
    popup.searchParams.set("email", email);
    const res = await harness.fetchWorker(TEST_GATEKEEPER_WORKER, popup.toString());
    if (res.status !== 200) {
      throw new Error(`Completing the sign-in flow failed with ${res.status}: ${await res.text()}`);
    }
    return await waiting;
  } finally {
    attempt[Symbol.dispose]();
  }
}

async function signInAndAuthenticate(
    publicApi: RpcStub<PublicApi>, email: string): Promise<RpcStub<AuthenticatedApi>> {
  const token = await signIn(publicApi, email);
  return (await publicApi.authenticate(token)) as unknown as RpcStub<AuthenticatedApi>;
}

/**
 * The bytes stored as this user's avatar, as a plain Uint8Array. Cap'n Web hands byte arrays to a
 * Node client as Buffers, and a Buffer never deep-equals the Uint8Array holding the same bytes.
 */
async function storedAvatar(
    api: RpcStub<AuthenticatedApi>, userId: string): Promise<Uint8Array | null> {
  const bytes = await api.getAvatar(userId);
  return bytes === null ? null : new Uint8Array(bytes);
}

describe("gatekeeper sign-in", () => {
  it.concurrent("offers the allowlisted auth gatekeeper as a sign-in method", async () => {
    await withSession(async publicApi => {
      const config = await publicApi.getServerConfig();
      expect(config.authVendors).toEqual([
        expect.objectContaining({ vendorId: TEST_VENDOR_ID, displayName: "Test Gatekeeper" }),
      ]);
      // The allowlist adds a sign-in button; it does not take the password form away.
      expect(config.passwordAuthEnabled).toBe(true);
    });
  });

  it.concurrent("signs a new user in and seeds their name and avatar from the gatekeeper",
      async () => {
    await withSession(async publicApi => {
      const email = signInEmail("newcomer");
      using api = await signInAndAuthenticate(publicApi, email);

      // The session is real: it authenticates, and it belongs to the email the gatekeeper verified.
      const me = await api.whoami();
      expect(me).toMatchObject({ type: "user", id: email, name: PROFILE_NAME });
      // The whole point of seeding: not the mangled email local-part the Workshop falls back to.
      expect(me.name).not.toBe(email.split("@")[0]);

      // And the photo made it from the gatekeeper into the avatar store, byte for byte.
      const hints = await fixtureProfileHints();
      expect(hints.name).toBe(PROFILE_NAME);
      await expect(storedAvatar(api, email)).resolves.toEqual(hints.photo);
    });
  });

  it.concurrent("leaves a name the user chose, and an avatar they set, alone on a later sign-in",
      async () => {
    await withSession(async publicApi => {
      const email = signInEmail("returning");
      {
        using first = await signInAndAuthenticate(publicApi, email);
        await first.setOwnDisplayName("Chosen Name");
        await first.setAvatar(CUSTOM_AVATAR);
      }

      // Signing in again offers the same hints. Neither may overwrite what the user settled on.
      using second = await signInAndAuthenticate(publicApi, email);
      await expect(second.whoami()).resolves.toMatchObject({ id: email, name: "Chosen Name" });
      await expect(storedAvatar(second, email)).resolves.toEqual(CUSTOM_AVATAR);
    });
  });
});
