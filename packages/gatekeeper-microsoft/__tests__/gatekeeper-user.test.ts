import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { GatekeeperUserImpl, GatekeeperVendor, UserAccount } from "../src/microsoft";

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
  /** Receives the request URL, so a test can answer each photo size differently. */
  photo?: (url: string) => Response;
};

/** Routes Graph paths to canned responses and records which paths were hit. */
function stubGraph(routes: GraphRoutes) {
  const calls: string[] = [];
  const handler = vi.fn(async (url: string) => {
    calls.push(url);
    // Both the sized renditions (/me/photos/<size>/$value) and the original (/me/photo/$value).
    if (url.includes("/me/photo")) {
      return routes.photo?.(url) ?? new Response(null, { status: 404 });
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

const PHOTO_240 = "https://graph.microsoft.com/v1.0/me/photos/240x240/$value";
const PHOTO_ORIGINAL = "https://graph.microsoft.com/v1.0/me/photo/$value";
const PHOTO_96 = "https://graph.microsoft.com/v1.0/me/photos/96x96/$value";

function photoResponse(byteLength: number, contentType = "image/jpeg"): Response {
  return new Response(new Uint8Array(byteLength).fill(7), {
    headers: { "Content-Type": contentType },
  });
}

function photoRequests(calls: string[]): string[] {
  return calls.filter(url => url.includes("/me/photo"));
}

describe("getAuthenticatedProfile", () => {
  it("returns the display name and a photo big enough for the Workshop's avatar", async () => {
    const { calls } = stubGraph({ photo: () => photoResponse(1024) });

    const profile = await user.getAuthenticatedProfile();

    expect(profile.name).toBe("Alex Morgan");
    expect(profile.photo?.mimeType).toBe("image/jpeg");
    expect(profile.photo?.data).toHaveLength(1024);
    // One rung: nothing further is asked for once a usable photo comes back.
    expect(photoRequests(calls)).toEqual([PHOTO_240]);
  });

  it("returns the name alone when the account has no photo", async () => {
    const { calls } = stubGraph({ photo: () => new Response(null, { status: 404 }) });

    // The key is absent, not undefined: the caller distinguishes "no hint" from "empty hint".
    await expect(user.getAuthenticatedProfile()).resolves.toStrictEqual({ name: "Alex Morgan" });
    expect(photoRequests(calls)).toEqual([PHOTO_240, PHOTO_ORIGINAL, PHOTO_96]);
  });

  it("drops a photo the Workshop would reject and keeps the name", async () => {
    // The Workshop stores these bytes as an avatar and accepts only JPEG and PNG, so a GIF or WebP
    // photo has to be dropped here rather than failing validation inside the sign-in callback.
    for (const contentType of ["image/webp", "image/gif", "image/svg+xml"]) {
      stubGraph({ photo: () => photoResponse(1024, contentType) });

      await expect(user.getAuthenticatedProfile()).resolves.toStrictEqual({ name: "Alex Morgan" });
    }
  });

  it("omits a blank display name instead of offering it as a hint", async () => {
    stubGraph({
      me: () => new Response(JSON.stringify({ ...MEMBER_PROFILE, displayName: "   " }), {
        headers: { "Content-Type": "application/json" },
      }),
      photo: () => photoResponse(1024),
    });

    const profile = await user.getAuthenticatedProfile();

    expect(profile.name).toBeUndefined();
    expect(Object.keys(profile)).toEqual(["photo"]);
  });

  it("falls back to a smaller rendition when the photo is over the size limit", async () => {
    // A Worker cannot re-encode an image, so an oversize photo is answered by asking Graph for a
    // smaller one.
    const { calls } = stubGraph({
      photo: url => url === PHOTO_96 ? photoResponse(4 * 1024) : photoResponse(200 * 1024),
    });

    const profile = await user.getAuthenticatedProfile();

    expect(profile.photo?.data).toHaveLength(4 * 1024);
    expect(photoRequests(calls)).toEqual([PHOTO_240, PHOTO_ORIGINAL, PHOTO_96]);
  });

  it("falls back to the original when no sized rendition is served", async () => {
    const { calls } = stubGraph({
      photo: url => url === PHOTO_ORIGINAL
        ? photoResponse(2048, "image/PNG; charset=binary")
        : new Response(null, { status: 404 }),
    });

    const profile = await user.getAuthenticatedProfile();

    expect(profile.photo?.mimeType).toBe("image/png");
    expect(photoRequests(calls)).toEqual([PHOTO_240, PHOTO_ORIGINAL]);
  });

  it("returns no hints once the transient grant has been cleaned up", async () => {
    stubGraph({ photo: () => photoResponse(1024) });
    await account.alarm();

    await expect(user.getAuthenticatedProfile()).resolves.toStrictEqual({});
  });

  it("returns no hints rather than throwing when Graph fails", async () => {
    // Sign-in must survive this, and a photo on its own is not worth a partial answer.
    stubGraph({
      me: () => new Response("boom", { status: 500 }),
      photo: () => photoResponse(1024),
    });

    await expect(user.getAuthenticatedProfile()).resolves.toStrictEqual({});
  });

  it("is advertised on the vendor, which is where the Workshop reads the capability", async () => {
    // An RPC stub cannot be probed for an optional method, so the flag is what makes the Workshop
    // call getAuthenticatedProfile() at all.
    const description = await new GatekeeperVendor({} as never, env as never).describe();

    expect(description.providesAuth).toBe(true);
    expect(description.providesAuthProfile).toBe(true);
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
