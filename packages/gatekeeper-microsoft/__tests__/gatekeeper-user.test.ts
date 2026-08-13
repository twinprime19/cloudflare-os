import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { GatekeeperUserImpl, UserAccount } from "../src/microsoft";

const TENANT = "11111111-2222-3333-4444-555555555555";
const DO_ID = "a".repeat(64);

const env = {
  CLIENT_ID: "client-id",
  CLIENT_SECRET: "client-secret",
  TENANT_ID: TENANT,
  BASE_URL: "https://gatekeeper.example/gatekeeper/microsoft",
};

const MEMBER_PROFILE = {
  id: "user-1",
  displayName: "Alex Morgan",
  mail: "Alex@Contoso.com",
  userPrincipalName: "alex@contoso.com",
  userType: "Member",
};

const ORGANIZATION = {
  value: [{ verifiedDomains: [{ name: "contoso.com" }, { name: "contoso.onmicrosoft.com" }] }],
};

function fakeDurableObjectContext() {
  const values = new Map<string, unknown>();
  return {
    id: { toString: () => DO_ID },
    storage: {
      setAlarm: vi.fn(),
      deleteAlarm: vi.fn(),
      deleteAll: vi.fn(() => values.clear()),
      kv: {
        get<T>(key: string) { return values.get(key) as T | undefined; },
        put<T>(key: string, value: T) { values.set(key, value); },
        delete(key: string) { values.delete(key); },
      },
    },
    exports: { GatekeeperUserImpl: vi.fn((init: unknown) => ({ userStub: init })) },
  };
}

type GraphRoutes = {
  me?: () => Response;
  organization?: () => Response;
  photo?: () => Response;
};

/** Routes Graph paths to canned responses and records which paths were hit. */
function stubGraph(routes: GraphRoutes) {
  const calls: string[] = [];
  const handler = vi.fn(async (url: string) => {
    calls.push(url);
    if (url.includes("/me/photos/")) {
      return routes.photo?.() ?? new Response(null, { status: 404 });
    }
    if (url.includes("/organization")) {
      return routes.organization?.()
        ?? new Response(JSON.stringify(ORGANIZATION), {
          headers: { "Content-Type": "application/json" },
        });
    }
    if (url.includes("/me")) {
      return routes.me?.()
        ?? new Response(JSON.stringify(MEMBER_PROFILE), {
          headers: { "Content-Type": "application/json" },
        });
    }
    throw new Error(`unexpected request: ${url}`);
  });
  vi.stubGlobal("fetch", handler);
  return { calls };
}

let context: ReturnType<typeof fakeDurableObjectContext>;
let account: UserAccount;
let user: GatekeeperUserImpl;

beforeEach(() => {
  context = fakeDurableObjectContext();
  account = new UserAccount(context as never, env as never);
  // A completed sign-in grant: an access token from the code exchange, and the claims recorded from
  // the id_token the token endpoint returned with it.
  context.storage.kv.put("accessToken", {
    token: "access-1", expires: new Date(Date.now() + 30 * 60 * 1000),
  });
  context.storage.kv.put("idTokenClaims", { tid: TENANT, oid: "object-1" });
  context.storage.kv.put("authOnly", true);

  const userContext = {
    props: { userObjectId: DO_ID },
    exports: {
      UserAccount: { idFromString: (id: string) => id, get: () => account },
      MicrosoftVerifier: vi.fn((init: unknown) => ({ verifierStub: init })),
    },
  };
  user = new GatekeeperUserImpl(userContext as never, env as never);
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("getAuthenticatedEmail", () => {
  it("returns the lowercased address of a member on a tenant-verified domain", async () => {
    stubGraph({});

    await expect(user.getAuthenticatedEmail()).resolves.toBe("alex@contoso.com");
  });

  it("caches the tenant's verified domains across sign-ins", async () => {
    const { calls } = stubGraph({});

    await user.getAuthenticatedEmail();
    await user.getAuthenticatedEmail();

    expect(calls.filter(url => url.includes("/organization"))).toHaveLength(1);
    expect(calls.filter(url => url.includes("/me?"))).toHaveLength(2);
  });

  it("returns null rather than throwing when Graph fails", async () => {
    stubGraph({ me: () => new Response("boom", { status: 500 }) });

    await expect(user.getAuthenticatedEmail()).resolves.toBeNull();
  });

  it("returns null once the transient grant has been cleaned up", async () => {
    stubGraph({});
    await account.alarm();

    await expect(user.getAuthenticatedEmail()).resolves.toBeNull();
  });

  it("returns null when the tenant reports no verified domains", async () => {
    stubGraph({
      organization: () => new Response(JSON.stringify({ value: [] }), {
        headers: { "Content-Type": "application/json" },
      }),
    });

    await expect(user.getAuthenticatedEmail()).resolves.toBeNull();
  });

  it("returns null when the deployment names no tenant", async () => {
    stubGraph({});
    const untenanted = new GatekeeperUserImpl(
      { props: { userObjectId: DO_ID }, exports: { UserAccount: { idFromString: (id: string) => id, get: () => account } } } as never,
      { ...env, TENANT_ID: undefined } as never);

    await expect(untenanted.getAuthenticatedEmail()).resolves.toBeNull();
  });
});

describe("describe", () => {
  it("uses the profile photo when there is one", async () => {
    stubGraph({
      photo: () => new Response(new Uint8Array([1, 2, 3]), {
        headers: { "Content-Type": "image/jpeg" },
      }),
    });

    const description = await user.describe();

    expect(description.displayName).toBe("Alex Morgan");
    expect(description.uniqueName).toBe("Alex@Contoso.com");
    expect(description.avatar.url).toBe("data:image/jpeg;base64,AQID");
  });

  it("falls back to the vendor logo when the account has no photo", async () => {
    stubGraph({ photo: () => new Response(null, { status: 404 }) });

    const description = await user.describe();

    expect(description.avatar.url).toContain("image/svg+xml");
  });
});

describe("resource surface", () => {
  it("offers no connectable resources yet", async () => {
    await expect(user.getSupportedResources()).resolves.toEqual([]);
    await expect(user.ensureResources([])).resolves.toEqual({});
    await expect(user.getGatekeeperClassFor("https://outlook.office.com/mail/"))
      .rejects.toThrow(/connectable resources/i);
    await expect(user.startResourceConfigurator("https://outlook.office.com/mail/*"))
      .rejects.toThrow(/Unsupported resource configurator/i);
    await expect(user.ensureResources(["https://outlook.office.com/mail/*"]))
      .rejects.toThrow(/Unknown grantable resource/i);
  });
});
