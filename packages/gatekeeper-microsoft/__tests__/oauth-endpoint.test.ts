import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { RpcStub } from "capnweb";

import worker, { GatekeeperVendor, UserAccount } from "../src/microsoft";

const TENANT = "11111111-2222-3333-4444-555555555555";
const BASE_URL = "https://gatekeeper.example/gatekeeper/microsoft";
const IDENTITY_SCOPES = ["openid", "profile", "email", "User.Read", "offline_access"];
const DO_ID = "a".repeat(64);
const INITIATION_NONCE = "n".repeat(64);

const env = {
  CLIENT_ID: "client-id",
  CLIENT_SECRET: "client-secret",
  TENANT_ID: TENANT,
  BASE_URL,
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

/** An ExecutionContext whose UserAccount namespace hands back one real Durable Object. */
function fakeExecutionContext(account: UserAccount) {
  return {
    exports: {
      UserAccount: {
        idFromString: (id: string) => ({ toString: () => id }),
        get: () => account,
      },
    },
  };
}

function get(path: string, search = ""): Request {
  return new Request(`${BASE_URL}${path}${search}`);
}

let account: UserAccount;
let context: ReturnType<typeof fakeDurableObjectContext>;
let executionContext: ReturnType<typeof fakeExecutionContext>;

beforeEach(() => {
  context = fakeDurableObjectContext();
  account = new UserAccount(context as never, env as never);
  executionContext = fakeExecutionContext(account);
  vi.stubGlobal("fetch", vi.fn(async () => new Response("{}", {
    headers: { "Content-Type": "application/json" },
  })));
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("authorization request", () => {
  it("redirects to the tenant's authorize endpoint without asking for a front-channel id_token",
     async () => {
    await account.setCallback({} as never, INITIATION_NONCE, IDENTITY_SCOPES, false);

    const response = await worker.fetch(
      get(`/${DO_ID}/${INITIATION_NONCE}`), env as never, executionContext as never);

    expect(response.status).toBe(302);
    const location = new URL(response.headers.get("Location")!);
    expect(location.origin + location.pathname)
      .toBe(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/authorize`);
    expect(location.searchParams.get("client_id")).toBe("client-id");
    expect(location.searchParams.get("redirect_uri")).toBe(`${BASE_URL}/oauth`);
    expect(location.searchParams.get("scope")).toBe(IDENTITY_SCOPES.join(" "));
    expect(location.searchParams.get("state")).toMatch(new RegExp(`^${DO_ID}:[0-9a-f]{64}$`));
    // Only an authorization code may cross the front channel: an id_token delivered here would be
    // attacker-supplied data, and the claims are read unverified on the assumption that they can
    // only have come from the token endpoint.
    expect(location.searchParams.get("response_type")).toBe("code");
    expect(location.searchParams.get("response_mode")).toBe("query");
  });

  it("shows the configuration page when the tenant is not set", async () => {
    const response = await worker.fetch(
      get(`/${DO_ID}/${INITIATION_NONCE}`),
      { ...env, TENANT_ID: undefined } as never,
      executionContext as never);

    expect(response.headers.get("Content-Type")).toBe("text/html; charset=utf-8");
    await expect(response.text()).resolves.toContain("Microsoft Gatekeeper Not Configured");
  });

  it("shows the expired-link page for an unknown nonce", async () => {
    await account.setCallback({} as never, INITIATION_NONCE, IDENTITY_SCOPES, false);

    const response = await worker.fetch(
      get(`/${DO_ID}/${"z".repeat(64)}`), env as never, executionContext as never);

    await expect(response.text()).resolves.toContain("Authorization Link Expired");
  });
});

describe("completion redirect", () => {
  it("serves provider errors as plain text so reflected markup cannot execute", async () => {
    const response = await worker.fetch(
      get("/oauth", "?error=access_denied&error_description=" +
        encodeURIComponent("<script>alert(document.domain)</script>")),
      env as never, executionContext as never);

    expect(response.status).toBe(400);
    expect(response.headers.get("Content-Type")).toBe("text/plain; charset=utf-8");
    expect(response.headers.get("X-Content-Type-Options")).toBe("nosniff");
    const body = await response.text();
    expect(body).toContain("access_denied");
    expect(body).toContain("<script>alert(document.domain)</script>");
  });

  it("caps the length of both echoed values", async () => {
    const response = await worker.fetch(
      get("/oauth", `?error=${"e".repeat(2000)}&error_description=${"d".repeat(2000)}`),
      env as never, executionContext as never);

    const body = await response.text();
    // Two 500-character values plus the separator and two ellipses.
    expect(body.length).toBeLessThan(1100);
    expect(body).toContain("…");
  });

  it("closes the popup once the code is redeemed", async () => {
    await account.setCallback({
      complete: async () => {}, credentialsExpired: async () => {},
      credentialsRestored: async () => {},
    } as never, INITIATION_NONCE, IDENTITY_SCOPES, false);
    const begun = await account.beginOAuthFlow(INITIATION_NONCE);
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({
      access_token: "access-1", expires_in: 3600, refresh_token: "refresh-1", scope: "openid",
    }), { headers: { "Content-Type": "application/json" } })));

    const response = await worker.fetch(
      get("/oauth", `?code=auth-code&state=${DO_ID}:${begun!.oauthNonce}`),
      env as never, executionContext as never);

    expect(response.headers.get("Content-Type")).toBe("text/html; charset=utf-8");
    await expect(response.text()).resolves.toContain("window.close()");
    expect(context.storage.kv.get("refreshToken")).toBe("refresh-1");
  });

  it("shows the expired-link page when the state nonce does not match", async () => {
    await account.setCallback({} as never, INITIATION_NONCE, IDENTITY_SCOPES, false);
    await account.beginOAuthFlow(INITIATION_NONCE);

    const response = await worker.fetch(
      get("/oauth", `?code=auth-code&state=${DO_ID}:${"z".repeat(64)}`),
      env as never, executionContext as never);

    await expect(response.text()).resolves.toContain("Authorization Link Expired");
  });

  it("rejects a missing or malformed state and a missing code", async () => {
    const cases: [string, string][] = [
      ["?code=auth-code", "no 'state' provided"],
      ["?code=auth-code&state=nocolon", "malformed state"],
      [`?state=${DO_ID}:${"z".repeat(64)}`, "no 'code' provided"],
    ];
    for (const [search, expected] of cases) {
      const response = await worker.fetch(
        get("/oauth", search), env as never, executionContext as never);
      await expect(response.text()).resolves.toContain(expected);
    }
  });
});

describe("routing", () => {
  it("404s an unknown path under the base URL", async () => {
    const response = await worker.fetch(
      get("/nope"), env as never, executionContext as never);
    expect(response.status).toBe(404);
  });

  it("refuses a request outside the configured base path", async () => {
    await expect(worker.fetch(
      new Request("https://gatekeeper.example/elsewhere"), env as never, executionContext as never))
      .rejects.toThrow(/does not match BASE_URL path/);
  });
});

describe("requested scopes", () => {
  /** A vendor whose UserAccount namespace mints one id and records what setCallback received. */
  function newVendor() {
    const setCallback = vi.fn(async () => {});
    const vendorContext = {
      exports: {
        UserAccount: {
          newUniqueId: () => ({ toString: () => DO_ID }),
          get: () => ({ setCallback }),
        },
      },
    };
    return {
      setCallback,
      vendor: new GatekeeperVendor(vendorContext as never, env as never),
    };
  }

  it("asks for the mailbox permission when connecting the mailbox resource", async () => {
    const { vendor, setCallback } = newVendor();

    const result = await vendor.connectAccount(new RpcStub({}) as never, {
      resourceUrlPatterns: ["https://outlook.office.com/mail/*"],
    });

    expect(result.url).toBe(`${BASE_URL}/${DO_ID}/${setCallback.mock.calls[0][1]}`);
    expect(setCallback.mock.calls[0][2]).toEqual([...IDENTITY_SCOPES, "Mail.ReadWrite"]);
  });

  it("asks for identity scopes only on a sign-in-only connection", async () => {
    const { vendor, setCallback } = newVendor();

    await vendor.connectAccount(new RpcStub({}) as never, { scopes: "auth" });

    expect(setCallback.mock.calls[0][2]).toEqual(
      IDENTITY_SCOPES.filter(scope => scope !== "offline_access"));
    expect(setCallback.mock.calls[0][3]).toBe(true);
  });

  it("refuses to connect an unknown resource pattern", async () => {
    const { vendor } = newVendor();

    await expect(vendor.connectAccount(new RpcStub({}) as never, {
      resourceUrlPatterns: ["https://example.com/*"],
    })).rejects.toThrow(/Unknown grantable resource/);
  });
});
