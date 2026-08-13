import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { UserAccount } from "../src/microsoft";

const TENANT = "11111111-2222-3333-4444-555555555555";
const TOKEN_URL = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;
const AUTH_SCOPES = ["openid", "profile", "email", "User.Read"];
const IDENTITY_SCOPES = [...AUTH_SCOPES, "offline_access"];

const env = {
  CLIENT_ID: "client-id",
  CLIENT_SECRET: "client-secret",
  TENANT_ID: TENANT,
  BASE_URL: "https://gatekeeper.example/gatekeeper/microsoft",
};

function base64url(value: unknown): string {
  return Buffer.from(JSON.stringify(value), "utf8").toString("base64url");
}

function idToken(claims: Record<string, unknown>): string {
  return `${base64url({ alg: "RS256" })}.${base64url(claims)}.signature`;
}

function fakeContext() {
  const values = new Map<string, unknown>();
  return {
    id: { toString: () => "a".repeat(64) },
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
    exports: {
      GatekeeperUserImpl: vi.fn((init: unknown) => ({ userStub: init })),
    },
  };
}

function fakeCallback() {
  return {
    complete: vi.fn(async () => {}),
    credentialsExpired: vi.fn(async () => {}),
    credentialsRestored: vi.fn(async () => {}),
  };
}

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status, headers: { "Content-Type": "application/json" },
  });
}

function newAccount(context = fakeContext()) {
  return { context, account: new UserAccount(context as never, env as never) };
}

/** Drive a full connect through the nonce lifecycle and return what the token endpoint saw. */
async function connect(account: UserAccount, callback: ReturnType<typeof fakeCallback>,
                       options: { authOnly?: boolean } = {}) {
  const initiationNonce = "n".repeat(64);
  const scopes = options.authOnly ? AUTH_SCOPES : IDENTITY_SCOPES;
  await account.setCallback(callback as never, initiationNonce, scopes, options.authOnly ?? false);
  const begun = await account.beginOAuthFlow(initiationNonce);
  expect(begun).not.toBeNull();
  return await account.acceptAuthCode("auth-code", begun!.oauthNonce);
}

/** Drive a reconnect through the same nonce lifecycle. */
async function reconnect(account: UserAccount, nonce = "r".repeat(64)) {
  await account.prepareReconnect(nonce, IDENTITY_SCOPES);
  const begun = await account.beginOAuthFlow(nonce);
  expect(begun).not.toBeNull();
  return await account.acceptAuthCode("auth-code-2", begun!.oauthNonce);
}

let fetchMock: ReturnType<typeof vi.fn>;

beforeEach(() => {
  fetchMock = vi.fn(async () => jsonResponse({}));
  vi.stubGlobal("fetch", fetchMock);
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function lastTokenRequestBody(): URLSearchParams {
  const call = fetchMock.mock.calls.at(-1) as [string, RequestInit];
  expect(call[0]).toBe(TOKEN_URL);
  return call[1].body as URLSearchParams;
}

describe("connect flow", () => {
  it("stores the grant and completes without reporting an expiry", async () => {
    const { context, account } = newAccount();
    const callback = fakeCallback();
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-1",
      expires_in: 3600,
      refresh_token: "refresh-1",
      scope: "openid profile email https://graph.microsoft.com/User.Read",
      id_token: idToken({ tid: TENANT, oid: "object-1" }),
    }));

    await expect(connect(account, callback)).resolves.toBe(true);

    const body = lastTokenRequestBody();
    expect(body.get("grant_type")).toBe("authorization_code");
    expect(body.get("code")).toBe("auth-code");
    expect(body.get("redirect_uri")).toBe(`${env.BASE_URL}/oauth`);
    expect(body.get("scope")).toBe(IDENTITY_SCOPES.join(" "));

    expect(context.storage.kv.get("refreshToken")).toBe("refresh-1");
    expect(context.storage.kv.get("idTokenClaims")).toEqual({ tid: TENANT, oid: "object-1" });
    await expect(account.getGrantScopes()).resolves.toEqual(IDENTITY_SCOPES);

    // `expiresAt` means "when the credentials stop being refreshable", which Entra never reports,
    // so complete() must be called with the user stub alone.
    expect(callback.complete).toHaveBeenCalledTimes(1);
    expect(callback.complete.mock.calls[0]).toHaveLength(1);
    expect(callback.credentialsRestored).not.toHaveBeenCalled();
  });

  it("refuses a persistent connection that came back with no refresh token", async () => {
    const { account } = newAccount();
    fetchMock.mockResolvedValue(jsonResponse({ access_token: "access-1", expires_in: 3600 }));

    await expect(connect(account, fakeCallback())).rejects.toThrow(/refresh token/i);
  });

  it("rejects a replayed or unknown nonce", async () => {
    const { account } = newAccount();
    const callback = fakeCallback();
    await account.setCallback(callback as never, "n".repeat(64), IDENTITY_SCOPES, false);

    expect(await account.beginOAuthFlow("m".repeat(64))).toBeNull();
    const begun = await account.beginOAuthFlow("n".repeat(64));
    expect(begun).not.toBeNull();
    // The initiation nonce is consumed by the first use.
    expect(await account.beginOAuthFlow("n".repeat(64))).toBeNull();

    expect(await account.acceptAuthCode("auth-code", "z".repeat(64))).toBe(false);
    expect(fetchMock).not.toHaveBeenCalled();

    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-1", expires_in: 3600, refresh_token: "refresh-1", scope: "openid",
    }));
    expect(await account.acceptAuthCode("auth-code", begun!.oauthNonce)).toBe(true);
    // The OAuth nonce is single-use too.
    expect(await account.acceptAuthCode("auth-code", begun!.oauthNonce)).toBe(false);
  });

  it("reports restoration rather than completion when the same account reconnects", async () => {
    const { context, account } = newAccount();
    const callback = fakeCallback();
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-1", expires_in: 3600, refresh_token: "refresh-1", scope: "openid",
      id_token: idToken({ tid: TENANT, oid: "object-1" }),
    }));
    await connect(account, callback);

    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-2", expires_in: 3600, refresh_token: "refresh-2", scope: "openid",
      id_token: idToken({ tid: TENANT, oid: "object-1" }),
    }));
    expect(await reconnect(account)).toBe(true);

    expect(context.storage.kv.get("refreshToken")).toBe("refresh-2");
    expect(callback.complete).toHaveBeenCalledTimes(1);
    expect(callback.credentialsRestored).toHaveBeenCalledTimes(1);
  });

  it("refuses a reconnect that signs in as a different Microsoft account", async () => {
    const { context, account } = newAccount();
    const callback = fakeCallback();
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-1", expires_in: 3600, refresh_token: "refresh-1", scope: "openid",
      id_token: idToken({ tid: TENANT, oid: "object-1" }),
    }));
    await connect(account, callback);

    // Same tenant, different directory principal: the consent screen lets the user pick any
    // account, and adopting it would silently repoint an established connection.
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-intruder", expires_in: 3600, refresh_token: "refresh-intruder",
      scope: "openid", id_token: idToken({ tid: TENANT, oid: "object-2" }),
    }));
    await expect(reconnect(account)).rejects.toThrow(/different Microsoft account/i);

    expect(context.storage.kv.get("refreshToken")).toBe("refresh-1");
    expect(context.storage.kv.get("idTokenClaims")).toEqual({ tid: TENANT, oid: "object-1" });
    expect(callback.credentialsRestored).not.toHaveBeenCalled();
    // The reconnect stays armed so the user can retry with the right account.
    expect(context.storage.kv.get("reconnecting")).toBe(true);
  });

  it("refuses a reconnect whose grant names no principal at all", async () => {
    const { context, account } = newAccount();
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-1", expires_in: 3600, refresh_token: "refresh-1", scope: "openid",
      id_token: idToken({ tid: TENANT, oid: "object-1" }),
    }));
    await connect(account, fakeCallback());

    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-2", expires_in: 3600, refresh_token: "refresh-2", scope: "openid",
    }));
    await expect(reconnect(account)).rejects.toThrow(/different Microsoft account/i);

    expect(context.storage.kv.get("refreshToken")).toBe("refresh-1");
  });

  it("refuses a reconnect on an account that never recorded a principal", async () => {
    // Entra returns `oid` on every openid-scoped grant, so an account with none is a state we
    // cannot verify — and an unverifiable reconnect is what this check exists to refuse.
    const { context, account } = newAccount();
    context.storage.kv.put("callback", fakeCallback());
    context.storage.kv.put("refreshToken", "refresh-1");
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-2", expires_in: 3600, refresh_token: "refresh-2", scope: "openid",
      id_token: idToken({ tid: TENANT, oid: "object-2" }),
    }));

    await expect(reconnect(account)).rejects.toThrow(/different Microsoft account/i);

    expect(context.storage.kv.get("refreshToken")).toBe("refresh-1");
    expect(context.storage.kv.get("idTokenClaims")).toBeUndefined();
    expect(context.storage.kv.get("reconnecting")).toBe(true);
  });

  it("discards the whole grant when the Workshop cannot accept the completed connection",
     async () => {
    const { context, account } = newAccount();
    const callback = fakeCallback();
    callback.complete.mockRejectedValue(new Error("workshop unreachable") as never);
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-1", expires_in: 3600, refresh_token: "refresh-1", scope: "openid",
      id_token: idToken({ tid: TENANT, oid: "object-1" }),
    }));

    await expect(connect(account, callback)).rejects.toThrow("workshop unreachable");

    // Nobody holds a usable account, so no part of the grant may survive — an access token left
    // behind is one a concurrent mint could still publish.
    expect(context.storage.kv.get("refreshToken")).toBeUndefined();
    expect(context.storage.kv.get("accessToken")).toBeUndefined();
    expect(context.storage.kv.get("idTokenClaims")).toBeUndefined();
    expect(context.storage.kv.get("grantScopes")).toBeUndefined();
  });
});

describe("refresh token rotation", () => {
  it("persists the rotated refresh token in place of the one it presented", async () => {
    const { context, account } = newAccount();
    context.storage.kv.put("refreshToken", "refresh-old");
    context.storage.kv.put("grantScopes", IDENTITY_SCOPES);
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-2",
      expires_in: 3600,
      refresh_token: "refresh-new",
      scope: "openid profile email https://graph.microsoft.com/User.Read offline_access",
    }));

    const token = await account.getAccessToken();

    expect(token.token).toBe("access-2");
    const body = lastTokenRequestBody();
    expect(body.get("grant_type")).toBe("refresh_token");
    expect(body.get("refresh_token")).toBe("refresh-old");
    expect(body.get("scope")).toBe(IDENTITY_SCOPES.join(" "));
    // Entra retires the presented token on every mint, so keeping the old value would strand the
    // account.
    expect(context.storage.kv.get("refreshToken")).toBe("refresh-new");
  });

  it("serves a cached token without minting, and re-mints once it is inside the safety window",
     async () => {
    const { context, account } = newAccount();
    context.storage.kv.put("refreshToken", "refresh-old");
    context.storage.kv.put("accessToken", {
      token: "access-cached", expires: new Date(Date.now() + 10 * 60 * 1000),
    });

    await expect(account.getAccessToken()).resolves.toMatchObject({ token: "access-cached" });
    expect(fetchMock).not.toHaveBeenCalled();

    context.storage.kv.put("accessToken", {
      token: "access-cached", expires: new Date(Date.now() + 5 * 1000),
    });
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-fresh", expires_in: 3600, refresh_token: "refresh-new", scope: "openid",
    }));
    await expect(account.getAccessToken()).resolves.toMatchObject({ token: "access-fresh" });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("collapses a burst of concurrent callers into a single token exchange", async () => {
    const { context, account } = newAccount();
    context.storage.kv.put("refreshToken", "refresh-old");
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-2", expires_in: 3600, refresh_token: "refresh-new", scope: "openid",
    }));

    const tokens = await Promise.all([
      account.getAccessToken(), account.getAccessToken(), account.getAccessToken(),
    ]);

    expect(tokens.map(token => token.token)).toEqual(["access-2", "access-2", "access-2"]);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    // Rotation happens inside the same locked section, so the racing callers cannot strand the
    // account by presenting or storing a retired refresh token.
    expect(context.storage.kv.get("refreshToken")).toBe("refresh-new");
  });
});

describe("mint failure taxonomy", () => {
  it("reports credential death once for a permanent rejection and stops asking", async () => {
    const { context, account } = newAccount();
    const callback = fakeCallback();
    context.storage.kv.put("callback", callback);
    context.storage.kv.put("refreshToken", "refresh-old");
    fetchMock.mockResolvedValue(jsonResponse({
      error: "invalid_grant",
      error_codes: [700082],
      error_description: "AADSTS700082: The refresh token has expired due to inactivity.",
    }, 400));

    await expect(account.getAccessToken()).rejects.toThrow(/reconnect the account/i);
    await expect(account.getAccessToken()).rejects.toThrow(/reconnect the account/i);

    expect(callback.credentialsExpired).toHaveBeenCalledTimes(1);
    // The second caller was answered from the recorded failure rather than another round trip.
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("keeps retrying a transient rejection and leaves the account alive", async () => {
    const { context, account } = newAccount();
    const callback = fakeCallback();
    context.storage.kv.put("callback", callback);
    context.storage.kv.put("refreshToken", "refresh-old");
    fetchMock.mockResolvedValue(jsonResponse({
      error: "temporarily_unavailable", error_codes: [90033],
    }, 503));

    await expect(account.getAccessToken()).rejects.toThrow(/try again/i);
    await expect(account.getAccessToken()).rejects.toThrow(/try again/i);

    expect(callback.credentialsExpired).not.toHaveBeenCalled();
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("clears a recorded failure when a reconnect supplies new credentials", async () => {
    const { context, account } = newAccount();
    const callback = fakeCallback();
    context.storage.kv.put("callback", callback);
    context.storage.kv.put("refreshToken", "refresh-old");
    context.storage.kv.put("idTokenClaims", { tid: TENANT, oid: "object-1" });
    fetchMock.mockResolvedValue(jsonResponse({ error_codes: [50173] }, 400));
    await expect(account.getAccessToken()).rejects.toThrow();
    expect(context.storage.kv.get("mintFailure")).toBeDefined();

    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-3", expires_in: 3600, refresh_token: "refresh-3", scope: "openid",
      id_token: idToken({ tid: TENANT, oid: "object-1" }),
    }));
    await reconnect(account);

    expect(context.storage.kv.get("mintFailure")).toBeUndefined();
    expect(callback.credentialsRestored).toHaveBeenCalledTimes(1);
  });

  it("treats a Graph claims challenge as credential death", async () => {
    const { context, account } = newAccount();
    const callback = fakeCallback();
    context.storage.kv.put("callback", callback);
    context.storage.kv.put("refreshToken", "refresh-old");
    context.storage.kv.put("accessToken", {
      token: "access-live", expires: new Date(Date.now() + 30 * 60 * 1000),
    });

    await account.reportCredentialsRejected("conditional access policy");

    expect(callback.credentialsExpired).toHaveBeenCalledTimes(1);
    // The rejected token must not keep being served from cache.
    expect(context.storage.kv.get("accessToken")).toBeUndefined();
    await expect(account.getAccessToken()).rejects.toThrow(/sign in again/i);
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe("sign-in-only grants", () => {
  it("is usable for the email read and then deletes itself", async () => {
    const { context, account } = newAccount();
    const callback = fakeCallback();
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-auth",
      expires_in: 3600,
      scope: "openid profile email https://graph.microsoft.com/User.Read",
      id_token: idToken({ tid: TENANT, oid: "object-1" }),
    }));

    await expect(connect(account, callback, { authOnly: true })).resolves.toBe(true);

    // No refresh token was requested, so none is stored...
    expect(lastTokenRequestBody().get("scope")).toBe(AUTH_SCOPES.join(" "));
    expect(context.storage.kv.get("refreshToken")).toBeUndefined();
    // ...yet the caller can still read the email with the access token from the exchange.
    await expect(account.getAccessToken()).resolves.toMatchObject({ token: "access-auth" });
    expect(fetchMock).toHaveBeenCalledTimes(1);

    const alarmAt = context.storage.setAlarm.mock.calls.at(-1)![0] as number;
    expect(alarmAt - Date.now()).toBeLessThanOrEqual(2 * 60 * 1000);

    await account.alarm();

    expect(context.storage.deleteAll).toHaveBeenCalled();
    expect(context.storage.kv.get("accessToken")).toBeUndefined();
    expect(context.storage.kv.get("idTokenClaims")).toBeUndefined();
  });

  it("deletes an abandoned flow that never produced credentials", async () => {
    const { context, account } = newAccount();
    await account.setCallback(fakeCallback() as never, "n".repeat(64), IDENTITY_SCOPES, false);

    await account.alarm();

    expect(context.storage.deleteAll).toHaveBeenCalled();
  });

  it("leaves a completed persistent connection alone when the abandonment alarm fires", async () => {
    const { context, account } = newAccount();
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-1", expires_in: 3600, refresh_token: "refresh-1", scope: "openid",
    }));
    await connect(account, fakeCallback());

    await account.alarm();

    expect(context.storage.deleteAll).not.toHaveBeenCalled();
    expect(context.storage.kv.get("refreshToken")).toBe("refresh-1");
  });
});

describe("revoke", () => {
  it("drops everything locally, since Entra offers no revocation endpoint for this flow",
     async () => {
    const { context, account } = newAccount();
    fetchMock.mockResolvedValue(jsonResponse({
      access_token: "access-1", expires_in: 3600, refresh_token: "refresh-1", scope: "openid",
    }));
    await connect(account, fakeCallback());
    fetchMock.mockClear();

    await account.revoke();

    expect(context.storage.deleteAlarm).toHaveBeenCalled();
    expect(context.storage.deleteAll).toHaveBeenCalled();
    expect(context.storage.kv.get("refreshToken")).toBeUndefined();
    expect(fetchMock).not.toHaveBeenCalled();
  });
});
