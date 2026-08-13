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
      OutlookMailGatekeeperImpl: vi.fn((init: unknown) => ({ gatekeeperStub: init })),
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

  it("reads the tenant's verified domains once per account, not once per lookup", async () => {
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

  it("tolerates media-type parameters on the photo response", async () => {
    stubGraph({
      photo: () => new Response(new Uint8Array([1, 2, 3]), {
        headers: { "Content-Type": "image/PNG; charset=binary" },
      }),
    });

    await expect(user.describe()).resolves.toMatchObject({
      avatar: { url: "data:image/png;base64,AQID" },
    });
  });

  it("falls back to the logo for a media type outside the image allowlist", async () => {
    // The type is copied into a URL the Workshop renders, so anything but a plain raster image is
    // treated exactly like "no photo" rather than passed through.
    for (const contentType of ["image/svg+xml", "text/html", "application/octet-stream", ""]) {
      stubGraph({
        photo: () => new Response(new Uint8Array([1, 2, 3]),
          contentType ? { headers: { "Content-Type": contentType } } : { headers: {} }),
      });

      const description = await user.describe();

      expect(description.avatar.url).toContain("image/svg+xml,");
      expect(description.avatar.url).not.toContain("base64");
    }
  });
});

const MAIL_PATTERN = "https://outlook.office.com/mail/*";

describe("resource surface", () => {
  it("offers the Outlook mailbox as the one grantable resource", async () => {
    const resources = await user.getSupportedResources();

    expect(resources).toHaveLength(1);
    expect(resources[0].urlPattern).toBe(MAIL_PATTERN);
    expect(resources[0].grantable).toBe(true);
  });

  it("routes a mailbox URL to the Outlook gatekeeper", async () => {
    const result = await user.getGatekeeperClassFor("https://outlook.office.com/mail/");

    expect(result.resource.urlPattern).toBe(MAIL_PATTERN);
    expect(result.class).toEqual({
      gatekeeperStub: { props: { userObjectId: DO_ID } },
    });
  });

  it("refuses a lookalike host and a non-mail path on the real host", async () => {
    await expect(user.getGatekeeperClassFor("https://outlook.office.com.evil.example/mail/"))
      .rejects.toThrow(/cannot connect this URL/i);
    await expect(user.getGatekeeperClassFor("https://evil.example/outlook.office.com/mail/"))
      .rejects.toThrow(/cannot connect this URL/i);
    await expect(user.getGatekeeperClassFor("https://outlook.office.com/calendar/view/month"))
      .rejects.toThrow(/cannot connect this URL/i);
    // The path must match on a segment boundary, not as a prefix.
    await expect(user.getGatekeeperClassFor("https://outlook.office.com/mailbox"))
      .rejects.toThrow(/cannot connect this URL/i);
    await expect(user.getGatekeeperClassFor("https://outlook.office.com/mailfoo/inbox"))
      .rejects.toThrow(/cannot connect this URL/i);
  });

  it("accepts the mailbox path with or without a trailing slash", async () => {
    await expect(user.getGatekeeperClassFor("https://outlook.office.com/mail")).resolves.toBeDefined();
    await expect(user.getGatekeeperClassFor("https://outlook.office.com/mail/inbox/id/123"))
      .resolves.toBeDefined();
  });

  it("serves the mailbox configurator frame", async () => {
    const frame = await user.startResourceConfigurator(MAIL_PATTERN);

    // The generated configurator HTML is a Text module the worker bundler inlines; under vitest the
    // import resolves to the module reference instead, so this asserts the wiring, not the markup.
    expect(typeof frame.iframeHtml).toBe("string");
    expect(frame.iframeHtml.length).toBeGreaterThan(0);
    expect(frame.ui).toBeDefined();
    await expect(user.startResourceConfigurator("https://example.com/*"))
      .rejects.toThrow(/Unsupported resource configurator/i);
  });

  it("asks for a reconnect only when the mailbox scope is missing", async () => {
    await expect(user.ensureResources([])).resolves.toEqual({});

    // No recorded scopes: the sign-in grant does not cover the mailbox.
    const expansion = await user.ensureResources([MAIL_PATTERN]);
    expect(expansion.url).toContain(`/gatekeeper/microsoft/${DO_ID}/`);

    // Entra reported the permission resource-qualified and differently cased; that still counts.
    context.storage.kv.put("grantedScopes", [
      "https://graph.microsoft.com/Mail.readwrite", "User.Read",
    ]);
    await expect(user.ensureResources([MAIL_PATTERN])).resolves.toEqual({});
  });

  it("rejects unknown resource patterns", async () => {
    await expect(user.ensureResources(["https://example.com/*"]))
      .rejects.toThrow(/Unknown grantable resource/i);
  });

  it("reports the granted resources on the account description", async () => {
    stubGraph({});
    context.storage.kv.put("grantedScopes", ["Mail.ReadWrite"]);

    const description = await user.describe();

    expect(description.grantedResourceUrlPatterns).toEqual([MAIL_PATTERN]);
  });
});
