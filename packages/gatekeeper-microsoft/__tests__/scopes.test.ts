import { RpcStub } from "capnweb";
import { describe, expect, it, vi } from "vitest";

import { grantCoversScopes, normalizeScope } from "../src/microsoft-api";
import { GatekeeperVendor, UserAccount } from "../src/microsoft";

describe("normalizeScope", () => {
  it("strips the Graph resource prefix and lowercases", () => {
    expect(normalizeScope("https://graph.microsoft.com/Mail.ReadWrite")).toBe("mail.readwrite");
    expect(normalizeScope("  Mail.Readwrite ")).toBe("mail.readwrite");
    expect(normalizeScope("User.Read")).toBe("user.read");
  });

  it("leaves a non-Graph scope alone apart from casing", () => {
    expect(normalizeScope("https://outlook.office.com/Mail.ReadWrite"))
      .toBe("https://outlook.office.com/mail.readwrite");
  });
});

describe("grantCoversScopes", () => {
  it("matches a resource-qualified response scope against a bare request scope", () => {
    expect(grantCoversScopes(
      ["https://graph.microsoft.com/Mail.ReadWrite", "https://graph.microsoft.com/User.Read"],
      ["Mail.ReadWrite"])).toBe(true);
  });

  it("ignores casing differences", () => {
    expect(grantCoversScopes(["mail.readwrite"], ["Mail.ReadWrite"])).toBe(true);
    expect(grantCoversScopes(["Mail.READWRITE"], ["Mail.ReadWrite"])).toBe(true);
  });

  it("treats the reserved OIDC scopes as granted even though Entra omits them", () => {
    expect(grantCoversScopes(
      ["https://graph.microsoft.com/User.Read"],
      ["openid", "profile", "email", "offline_access", "User.Read"])).toBe(true);
  });

  it("reports a genuinely missing permission as missing", () => {
    expect(grantCoversScopes(["https://graph.microsoft.com/User.Read"], ["Mail.ReadWrite"]))
      .toBe(false);
    expect(grantCoversScopes([], ["Mail.ReadWrite"])).toBe(false);
    // A read-only grant does not satisfy a read-write requirement.
    expect(grantCoversScopes(["Mail.Read"], ["Mail.ReadWrite"])).toBe(false);
  });
});

// ── Resource scope plumbing ─────────────────────────────────────────
//
// The mapping from grantable resources to OAuth scopes is exercised through the flows that use it —
// `connectAccount`, which is what asks Entra for them, and the account's own report of what the
// grant covers — rather than by exporting the internal helpers.

const DO_ID = "a".repeat(64);
const MAIL_PATTERN = "https://outlook.office.com/mail/*";
const TEAMS_PATTERN = "https://teams.microsoft.com/*";

const IDENTITY_SCOPES = ["openid", "profile", "email", "User.Read", "offline_access"];
const MAIL_SCOPES = ["Mail.ReadWrite"];
const TEAMS_SCOPES = [
  "Team.ReadBasic.All", "Channel.ReadBasic.All", "TeamMember.Read.All", "ChannelMessage.Read.All",
  "Chat.Read",
];

const env = {
  CLIENT_ID: "client-id",
  CLIENT_SECRET: "client-secret",
  TENANT_ID: "11111111-2222-3333-4444-555555555555",
  BASE_URL: "https://gatekeeper.example/gatekeeper/microsoft",
};

/** Captures the scope list `connectAccount` hands the account to request from Entra. */
function fakeVendor() {
  const requested: string[][] = [];
  const ctx = {
    exports: {
      UserAccount: {
        newUniqueId: () => ({ toString: () => DO_ID }),
        get: () => ({
          setCallback: async (
              _callback: unknown, _nonce: string, scopes: string[], _authOnly?: boolean) => {
            requested.push(scopes);
          },
        }),
      },
    },
  };
  return { vendor: new GatekeeperVendor(ctx as never, env as never), requested };
}

/** A real capnweb stub: the RPC argument validator refuses a plain object for the callback. */
function connectCallback(): never {
  return new RpcStub({}) as never;
}

async function scopesFor(options?: Record<string, unknown>): Promise<string[]> {
  const { vendor, requested } = fakeVendor();
  await vendor.connectAccount(connectCallback(), options as never);
  return requested[0];
}

/** An account whose grant reported exactly `grantedScopes`. */
function accountWithGrant(grantedScopes: string[]): UserAccount {
  const values = new Map<string, unknown>([["grantedScopes", grantedScopes]]);
  const ctx = {
    id: { toString: () => DO_ID },
    storage: {
      setAlarm: vi.fn(),
      deleteAlarm: vi.fn(),
      deleteAll: vi.fn(),
      kv: {
        get<T>(key: string) { return values.get(key) as T | undefined; },
        put<T>(key: string, value: T) { values.set(key, value); },
        delete(key: string) { values.delete(key); },
      },
    },
  };
  return new UserAccount(ctx as never, env as never);
}

describe("resourceUrlPatternsToOAuthScopes", () => {
  it("requests exactly the Teams scopes for a Teams-only connection", async () => {
    const scopes = await scopesFor({ resourceUrlPatterns: [TEAMS_PATTERN] });

    expect(scopes.toSorted()).toEqual([...IDENTITY_SCOPES, ...TEAMS_SCOPES].toSorted());
    expect(scopes).not.toContain("Mail.ReadWrite");
  });

  it("requests exactly the mail scope for a mail-only connection", async () => {
    const scopes = await scopesFor({ resourceUrlPatterns: [MAIL_PATTERN] });

    expect(scopes.toSorted()).toEqual([...IDENTITY_SCOPES, ...MAIL_SCOPES].toSorted());
    for (const scope of TEAMS_SCOPES) expect(scopes).not.toContain(scope);
  });

  it("requests mail AND Teams when no patterns are named", async () => {
    // Omitting the patterns means "every resource this vendor offers", and the generic Connect
    // Account flow omits them. So adding Teams widened what a plain connect asks for: this is that
    // behavior, pinned deliberately rather than discovered later from a consent screen.
    const scopes = await scopesFor(undefined);

    expect(scopes.toSorted())
      .toEqual([...IDENTITY_SCOPES, ...MAIL_SCOPES, ...TEAMS_SCOPES].toSorted());
  });

  it("requests no resource scopes for an empty pattern list", async () => {
    // An empty array is meaningfully different from omitting it: it asks for no data access.
    const scopes = await scopesFor({ resourceUrlPatterns: [] });

    expect(scopes.toSorted()).toEqual([...IDENTITY_SCOPES].toSorted());
  });

  it("leaves sign-in insulated from the resource scopes entirely", async () => {
    const scopes = await scopesFor({ scopes: "auth" });

    expect(scopes.toSorted()).toEqual(["openid", "profile", "email", "User.Read"].toSorted());
  });

  it("refuses a pattern this vendor does not offer", async () => {
    const { vendor } = fakeVendor();

    await expect(vendor.connectAccount(
        connectCallback(), { resourceUrlPatterns: ["https://example.com/*"] } as never))
      .rejects.toThrow(/Unknown grantable resource/i);
  });
});

describe("grantedResourcesFromScopes", () => {
  it("detects a Teams-only grant", async () => {
    await expect(accountWithGrant([...IDENTITY_SCOPES, ...TEAMS_SCOPES])
      .getGrantedResourceUrlPatterns()).resolves.toEqual([TEAMS_PATTERN]);
  });

  it("detects a mail-only grant", async () => {
    await expect(accountWithGrant([...IDENTITY_SCOPES, ...MAIL_SCOPES])
      .getGrantedResourceUrlPatterns()).resolves.toEqual([MAIL_PATTERN]);
  });

  it("detects a grant covering both", async () => {
    await expect(accountWithGrant([...IDENTITY_SCOPES, ...MAIL_SCOPES, ...TEAMS_SCOPES])
      .getGrantedResourceUrlPatterns()).resolves.toEqual([MAIL_PATTERN, TEAMS_PATTERN]);
  });

  it("treats a partly consented Teams grant as not granted at all", async () => {
    // Detection is all-or-nothing: one permission an administrator declined leaves the whole
    // resource unavailable, which is why the shortfall is logged when the grant lands.
    const partial = TEAMS_SCOPES.filter(scope => scope !== "ChannelMessage.Read.All");

    await expect(accountWithGrant([...IDENTITY_SCOPES, ...partial])
      .getGrantedResourceUrlPatterns()).resolves.toEqual([]);
  });

  it("accepts the resource-qualified, differently cased form Entra actually returns", async () => {
    const qualified = TEAMS_SCOPES.map(
        scope => `https://graph.microsoft.com/${scope.toLowerCase()}`);

    await expect(accountWithGrant(qualified).getGrantedResourceUrlPatterns())
      .resolves.toEqual([TEAMS_PATTERN]);
  });
});
