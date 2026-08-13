import { afterEach, describe, expect, it, vi } from "vitest";

import { GraphApiError, GraphMailApi, assertGraphUrl, graphUrl, validateSearchQuery } from "../src/graph-api";

type Call = { url: string; init: RequestInit };

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status, headers: { "Content-Type": "application/json", ...headers },
  });
}

/** Route every fetch through `handler`, recording what the client sent. */
function stubFetch(handler: (call: Call) => Response | Promise<Response>): Call[] {
  const calls: Call[] = [];
  vi.stubGlobal("fetch", vi.fn(async (url: string, init: RequestInit = {}) => {
    calls.push({ url, init });
    return await handler({ url, init });
  }));
  return calls;
}

function prefer(call: Call): string {
  return new Headers(call.init.headers).get("Prefer") ?? "";
}

function newApi(opts: { onCredentialsRejected?: (detail: string) => Promise<void> } = {}) {
  return new GraphMailApi(async () => "access-token", opts);
}

const MESSAGE = {
  id: "AAMkImmutable1",
  subject: "Quarterly report",
  from: { emailAddress: { address: "bob@example.com", name: "Bob" } },
  toRecipients: [{ emailAddress: { address: "me@example.com" } }],
  ccRecipients: [],
  receivedDateTime: "2026-08-01T10:00:00Z",
  isRead: false,
  hasAttachments: true,
  bodyPreview: "Numbers attached",
  parentFolderId: "folder-inbox",
  webLink: "https://outlook.office.com/mail/deeplink",
  conversationId: "conv-1",
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("URL construction", () => {
  it("percent-encodes every path segment", () => {
    const url = graphUrl(["me", "messages", "id/with?weird#chars", "move"]);

    expect(url).toBe(
      "https://graph.microsoft.com/v1.0/me/messages/id%2Fwith%3Fweird%23chars/move");
  });

  it("encodes query values instead of interpolating them", () => {
    const url = graphUrl(["me", "messages"], {
      $search: "\"from:bob&$top=999\"", $top: "25",
    });

    // The system option names stay literal; the value cannot start a second parameter.
    expect(url).toBe(
      "https://graph.microsoft.com/v1.0/me/messages" +
      "?$search=%22from%3Abob%26%24top%3D999%22&$top=25");
    expect(new URL(url).searchParams.get("$top")).toBe("25");
  });

  it("keeps an agent-supplied id inside one path segment", async () => {
    const calls = stubFetch(() => jsonResponse(MESSAGE));
    const hostileId = "../../users/victim@example.com/messages/stolen";

    await newApi().getMessage(hostileId);

    expect(calls[0].url).toBe(
      `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(hostileId)}` +
      `?$select=${encodeURIComponent(
        "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead," +
        "hasAttachments,bodyPreview,parentFolderId,webLink,conversationId")}`);
    // /v1.0/me/messages/<id> — the id never becomes extra path segments.
    expect(new URL(calls[0].url).pathname.split("/")).toHaveLength(5);
  });

  it("refuses a dot or empty segment, which URL parsing would collapse", () => {
    // `/v1.0/me/mailFolders/..` would normalize to `/v1.0/me` — a different endpoint, reached with
    // this mailbox's bearer token.
    expect(() => graphUrl(["me", "mailFolders", ".."])).toThrow(/empty or relative/);
    expect(() => graphUrl(["me", "mailFolders", "."])).toThrow(/empty or relative/);
    expect(() => graphUrl(["me", "mailFolders", ""])).toThrow(/empty or relative/);
    // A dot inside an id is still an ordinary character.
    expect(graphUrl(["me", "mailFolders", "a..b"]))
      .toBe("https://graph.microsoft.com/v1.0/me/mailFolders/a..b");
  });

  it("rejects a relative folder id before any request is made", async () => {
    const calls = stubFetch(() => jsonResponse({}));

    // `getFolder` is the only place an agent-supplied id reaches a path (message ids come from
    // Graph itself); a move's destination travels in the JSON body, where it cannot retarget a URL.
    await expect(newApi().getFolder("..")).rejects.toThrow(/empty or relative/);

    expect(calls).toHaveLength(0);
  });
});

describe("pagination link pinning", () => {
  it("accepts a Graph link", () => {
    expect(assertGraphUrl("https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc"))
      .toContain("graph.microsoft.com");
  });

  it("rejects a link on any other origin", () => {
    expect(() => assertGraphUrl("https://graph.microsoft.com.evil.example/v1.0/me/messages"))
      .toThrow(/outside https:\/\/graph.microsoft.com/);
    expect(() => assertGraphUrl("http://graph.microsoft.com/v1.0/me/messages"))
      .toThrow(/outside https:\/\/graph.microsoft.com/);
    expect(() => assertGraphUrl("not a url")).toThrow(/malformed pagination link/);
  });

  it("refuses to follow an off-origin nextLink returned by a listing", async () => {
    stubFetch(() => jsonResponse({
      value: [MESSAGE],
      "@odata.nextLink": "https://attacker.example/v1.0/me/messages?$skiptoken=abc",
    }));

    await expect(newApi().listMessages()).rejects.toThrow(/outside https:\/\/graph.microsoft.com/);
  });

  it("follows a Graph nextLink", async () => {
    const calls = stubFetch(call => call.url.includes("skiptoken")
      ? jsonResponse({ value: [{ ...MESSAGE, id: "AAMkImmutable2" }] })
      : jsonResponse({
        value: [MESSAGE],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc",
      }));

    const first = await newApi().listMessages();
    expect(first.nextLink).toBeDefined();
    const second = await newApi().nextMessagePage(first.nextLink!);

    expect(second.items[0].id).toBe("AAMkImmutable2");
    expect(calls).toHaveLength(2);
  });

  it("stops walking folders at the page cap", async () => {
    let page = 0;
    const calls = stubFetch(() => jsonResponse({
      value: [{ id: `folder-${page++}`, displayName: "Folder" }],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders?$skiptoken=next",
    }));

    const folders = await newApi().listFolders();

    expect(folders).toHaveLength(10);
    expect(calls).toHaveLength(10);
  });
});

describe("folder traversal", () => {
  it("descends into nested folders, which the root listing omits", async () => {
    // /me/mailFolders returns only the root's children, so "Clients/Acme" is reachable only via
    // childFolders.
    const calls = stubFetch(call => {
      const path = new URL(call.url).pathname;
      if (path.endsWith("/me/mailFolders")) {
        return jsonResponse({ value: [
          { id: "inbox", displayName: "Inbox", childFolderCount: 0 },
          { id: "clients", displayName: "Clients", childFolderCount: 2 },
        ] });
      }
      if (path.endsWith("/mailFolders/clients/childFolders")) {
        return jsonResponse({ value: [
          { id: "acme", displayName: "Acme", parentFolderId: "clients", childFolderCount: 1 },
          { id: "globex", displayName: "Globex", parentFolderId: "clients", childFolderCount: 0 },
        ] });
      }
      if (path.endsWith("/mailFolders/acme/childFolders")) {
        return jsonResponse({ value: [
          { id: "acme-2026", displayName: "2026", parentFolderId: "acme", childFolderCount: 0 },
        ] });
      }
      throw new Error(`unexpected request: ${call.url}`);
    });

    const folders = await newApi().listFolders();

    // Breadth-first: shallow folders first.
    expect(folders.map(folder => folder.name))
      .toEqual(["Inbox", "Clients", "Acme", "Globex", "2026"]);
    expect(folders.find(folder => folder.id === "acme")!.parentFolderId).toBe("clients");
    // Leaves are never visited: 1 root + 2 parents with children, not 1 + 5.
    expect(calls).toHaveLength(3);
  });

  it("costs exactly one request for a mailbox with no subfolders", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [
      { id: "inbox", displayName: "Inbox", childFolderCount: 0 },
      { id: "archive", displayName: "Archive" },
    ] }));

    expect(await newApi().listFolders()).toHaveLength(2);
    expect(calls).toHaveLength(1);
  });

  it("bounds a pathologically nested mailbox by total requests", async () => {
    let created = 0;
    const calls = stubFetch(() => jsonResponse({
      value: [{ id: `folder-${created++}`, displayName: "Folder", childFolderCount: 1 }],
    }));

    const folders = await newApi().listFolders();

    expect(calls).toHaveLength(25);
    expect(folders).toHaveLength(25);
  });
});

describe("request headers", () => {
  it("sends the immutable-id preference on every request", async () => {
    // One body that satisfies every read shape: a collection and a single entity.
    const calls = stubFetch(() => jsonResponse({ ...MESSAGE, displayName: "Inbox", value: [MESSAGE] }));
    const api = newApi();

    await api.listMessages();
    await api.listMessages({ folderId: "folder-inbox" });
    await api.searchMessages("from:bob@example.com");
    await api.getMessage(MESSAGE.id);
    await api.getMessageBody(MESSAGE.id);
    await api.listFolders();
    await api.getFolder("folder-inbox");
    await api.setMessageRead(MESSAGE.id, true);
    await api.moveMessage(MESSAGE.id, "folder-archive");
    await api.createReplyDraft(MESSAGE.id, "thanks", false);
    await api.createReplyDraft(MESSAGE.id, "thanks", true);

    expect(calls.length).toBe(11);
    for (const call of calls) {
      expect(prefer(call)).toContain('IdType="ImmutableId"');
    }
  });

  it("asks for plain-text bodies on reads only", async () => {
    const calls = stubFetch(() => jsonResponse(MESSAGE));
    const api = newApi();

    await api.getMessage(MESSAGE.id);
    await api.setMessageRead(MESSAGE.id, true);

    expect(prefer(calls[0])).toContain('outlook.body-content-type="text"');
    expect(prefer(calls[1])).not.toContain("outlook.body-content-type");
    expect(calls[1].init.method).toBe("PATCH");
  });
});

describe("reads", () => {
  it("maps a message onto the agent-facing shape", async () => {
    stubFetch(() => jsonResponse({ value: [MESSAGE] }));

    const page = await newApi().listMessages();

    expect(page.items[0]).toEqual({
      id: "AAMkImmutable1",
      subject: "Quarterly report",
      from: { address: "bob@example.com", name: "Bob" },
      to: [{ address: "me@example.com" }],
      cc: [],
      receivedAt: new Date("2026-08-01T10:00:00Z"),
      isRead: false,
      hasAttachments: true,
      preview: "Numbers attached",
      folderId: "folder-inbox",
      webLink: "https://outlook.office.com/mail/deeplink",
      conversationId: "conv-1",
    });
  });

  it("orders listings newest-first but never orders a search", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [] }));
    const api = newApi();

    await api.listMessages();
    await api.searchMessages("invoice");

    expect(calls[0].url).toContain("$orderby=receivedDateTime%20desc");
    expect(calls[1].url).not.toContain("orderby");
    expect(calls[1].url).toContain("$search=%22invoice%22");
  });

  it("rejects a search query that would break out of the quoted KQL string", () => {
    expect(() => validateSearchQuery("subject:\"x\" OR from:ceo@example.com"))
      .toThrow(/double quotes/);
    expect(() => validateSearchQuery("   ")).toThrow(/must not be empty/);
    expect(() => validateSearchQuery("a".repeat(401))).toThrow(/at most 400/);
    expect(validateSearchQuery("  from:bob  ")).toBe("from:bob");
  });
});

describe("error taxonomy", () => {
  it("retries a plain 401 once with a fresh token", async () => {
    const tokens = ["stale", "fresh"];
    let issued = 0;
    const calls = stubFetch(call =>
      new Headers(call.init.headers).get("Authorization") === "Bearer fresh"
        ? jsonResponse({ value: [MESSAGE] })
        : jsonResponse({ error: { code: "InvalidAuthenticationToken" } }, 401));

    const api = new GraphMailApi(async () => tokens[Math.min(issued++, 1)]);
    const page = await api.listMessages();

    expect(page.items).toHaveLength(1);
    expect(calls).toHaveLength(2);
  });

  it("reports a claims-challenge 401 as credential death and does not retry", async () => {
    const rejected: string[] = [];
    const calls = stubFetch(() => jsonResponse(
      { error: { code: "InvalidAuthenticationToken", message: "policy" } }, 401, {
        "WWW-Authenticate":
          'Bearer realm="", error="insufficient_claims", claims="eyJhY2Nlc3NfdG9rZW4ifQ=="',
      }));

    const api = newApi({
      onCredentialsRejected: async detail => { rejected.push(detail); },
    });

    const error = await api.listMessages().catch(err => err);

    expect(error).toBeInstanceOf(GraphApiError);
    expect((error as GraphApiError).credentialsRejected).toBe(true);
    expect((error as GraphApiError).message).toMatch(/sign in again/i);
    expect(rejected).toEqual(["insufficient_claims"]);
    // One attempt only: another token carries the same claims.
    expect(calls).toHaveLength(1);
  });

  it("surfaces a 403 as a permission problem", async () => {
    stubFetch(() => jsonResponse(
      { error: { code: "ErrorAccessDenied", message: "Access is denied." } }, 403));

    const error = await newApi().listMessages().catch(err => err) as GraphApiError;

    expect(error.status).toBe(403);
    expect(error.code).toBe("ErrorAccessDenied");
    expect(error.message).toMatch(/not permitted/i);
    expect(error.credentialsRejected).toBe(false);
  });

  it("honors Retry-After on a 429 and gives up with a throttling error", async () => {
    const calls = stubFetch(() => jsonResponse(
      { error: { code: "ApplicationThrottled", message: "slow down" } }, 429,
      { "Retry-After": "0" }));

    const error = await newApi().listMessages().catch(err => err) as GraphApiError;

    expect(error.status).toBe(429);
    expect(error.message).toMatch(/throttling/i);
    expect(calls).toHaveLength(3);
  });

  it("replays a throttled write, which Graph rejects before applying it", async () => {
    let attempts = 0;
    stubFetch(() => ++attempts === 1
      ? jsonResponse({ error: { code: "ApplicationThrottled" } }, 429, { "Retry-After": "0" })
      : new Response(null, { status: 204 }));

    await newApi().setMessageRead(MESSAGE.id, true);

    expect(attempts).toBe(2);
  });

  it("explains a 404 in mailbox terms", async () => {
    stubFetch(() => jsonResponse(
      { error: { code: "ErrorItemNotFound", message: "The specified object was not found." } },
      404));

    const error = await newApi().getMessage(MESSAGE.id).catch(err => err) as GraphApiError;

    expect(error.status).toBe(404);
    expect(error.code).toBe("ErrorItemNotFound");
    expect(error.message).toMatch(/no longer exists/i);
  });
});

describe("writes", () => {
  it("builds the documented mailbox write requests", async () => {
    const calls = stubFetch(() => new Response(null, { status: 204 }));
    const api = newApi();

    await api.setMessageRead(MESSAGE.id, false);
    await api.moveMessage(MESSAGE.id, "folder-archive");
    await api.createReplyDraft(MESSAGE.id, "on it", true);

    expect(calls[0].init.method).toBe("PATCH");
    expect(calls[0].init.body).toBe(JSON.stringify({ isRead: false }));
    expect(calls[1].url).toMatch(/\/me\/messages\/AAMkImmutable1\/move$/);
    expect(calls[1].init.body).toBe(JSON.stringify({ destinationId: "folder-archive" }));
    expect(calls[2].url).toMatch(/\/me\/messages\/AAMkImmutable1\/createReplyAll$/);
    expect(calls[2].init.body).toBe(JSON.stringify({ comment: "on it" }));
  });

  it("refuses an oversized reply body before sending anything", async () => {
    const calls = stubFetch(() => new Response(null, { status: 204 }));

    await expect(newApi().createReplyDraft(MESSAGE.id, "x".repeat(64 * 1024 + 1), false))
      .rejects.toThrow(/at most/);

    expect(calls).toHaveLength(0);
  });
});
