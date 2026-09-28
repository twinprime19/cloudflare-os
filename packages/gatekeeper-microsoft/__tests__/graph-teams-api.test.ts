import { afterEach, describe, expect, it, vi } from "vitest";

import { GraphApiError } from "../src/graph-api";
import {
  GraphChatMessage, GraphTeamsApi, buildSearchRequestBody, renderTeamsMessageBody,
  validateTeamsSearchQuery,
} from "../src/graph-teams-api";
import type { TeamsSearchHitInfo } from "../src/teams-types";

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

function newApi(opts: { onCredentialsRejected?: (detail: string) => Promise<void> } = {}) {
  return new GraphTeamsApi(async () => "access-token", opts);
}

function html(content: string): GraphChatMessage["body"] {
  return { contentType: "html", content };
}

const CHANNEL_MESSAGE: GraphChatMessage = {
  id: "1700000000000",
  messageType: "message",
  createdDateTime: "2026-08-01T10:00:00Z",
  lastEditedDateTime: "2026-08-01T10:05:00Z",
  webUrl: "https://teams.microsoft.com/l/message/19:channel/1700000000000",
  from: { user: { id: "user-1", displayName: "Bob" }, application: null, device: null },
  body: html("<div>Numbers attached</div>"),
  attachments: [],
  mentions: [],
};

/** Graph fills the members of an identity set that do not apply with explicit nulls. */
const BOT_IDENTITY: GraphChatMessage["from"] = {
  application: { id: "app-1", displayName: "Builds" },
  device: null,
  user: null,
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("URL construction", () => {
  it("keeps an agent-supplied id inside one path segment", async () => {
    const calls = stubFetch(() => jsonResponse(CHANNEL_MESSAGE));
    const hostileChat = "../../users/victim@example.com/chats/19:stolen";
    const hostileMessage = "1700000000000?$select=body";

    await newApi().getChatMessage(hostileChat, hostileMessage);

    expect(calls[0].url).toBe(
      "https://graph.microsoft.com/v1.0/chats/" +
      `${encodeURIComponent(hostileChat)}/messages/${encodeURIComponent(hostileMessage)}`);
    // /v1.0/chats/<id>/messages/<id> — neither id becomes extra path segments or a parameter.
    expect(new URL(calls[0].url).pathname.split("/")).toHaveLength(6);
    expect(new URL(calls[0].url).search).toBe("");
  });

  it("refuses a relative or empty id before any request is made", async () => {
    const calls = stubFetch(() => jsonResponse({}));
    const api = newApi();

    await expect(api.getTeam("..")).rejects.toThrow(/empty or relative/);
    await expect(api.getChannel("team-1", ".")).rejects.toThrow(/empty or relative/);
    await expect(api.listChatMembers("")).rejects.toThrow(/empty or relative/);
    await expect(api.listReplies("team-1", "channel-1", "..")).rejects.toThrow(/empty or relative/);

    expect(calls).toHaveLength(0);
  });
});

describe("endpoint shapes", () => {
  it("asks each collection for what Graph actually supports", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [] }));
    const api = newApi();

    await api.listChannelMessages("team-1", "channel-1");
    await api.listChatMessages("chat-1");
    await api.listChats();
    await api.listChatMembers("chat-1");
    await api.listTeamMembers("team-1");
    await api.listReplies("team-1", "channel-1", "message-1");

    const [channelMessages, chatMessages, chats, chatMembers, teamMembers, replies] =
      calls.map(call => new URL(call.url));

    // A channel is sorted by its reply chains' last-modified time and refuses any other order.
    expect(channelMessages.searchParams.get("$orderby")).toBeNull();
    expect(channelMessages.searchParams.get("$top")).toBe("25");
    // A chat is ordered explicitly, because Graph's default (last-modified) makes an edited old
    // message resurface at the front of a page walk.
    expect(chatMessages.searchParams.get("$orderby")).toBe("createdDateTime desc");
    // The members expansion truncates rosters at 25 per chat regardless of $top.
    expect(chats.searchParams.get("$expand")).toBeNull();
    expect(chats.searchParams.get("$top")).toBe("50");
    // Chat members reject OData parameters outright, so the page size is Graph's.
    expect(chatMembers.search).toBe("");
    expect(teamMembers.searchParams.get("$top")).toBe("50");
    expect(replies.searchParams.get("$top")).toBe("25");
    expect(replies.pathname).toMatch(/\/messages\/message-1\/replies$/);

    for (const call of calls) {
      const top = new URL(call.url).searchParams.get("$top");
      if (top !== null) expect(Number(top)).toBeLessThanOrEqual(50);
    }
  });

  it("sends no Outlook-only request preferences", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [] }));

    await newApi().listChats();

    expect(new Headers(calls[0].init.headers).get("Prefer")).toBeNull();
    expect(new Headers(calls[0].init.headers).get("Authorization")).toBe("Bearer access-token");
  });
});

describe("pagination", () => {
  it("follows a Graph nextLink verbatim", async () => {
    const nextLink =
      "https://graph.microsoft.com/v1.0/chats/chat-1/messages?$top=25&$skiptoken=abc";
    const calls = stubFetch(call => call.url.includes("skiptoken")
      ? jsonResponse({ value: [{ ...CHANNEL_MESSAGE, id: "1700000000001" }] })
      : jsonResponse({ value: [CHANNEL_MESSAGE], "@odata.nextLink": nextLink }));

    const api = newApi();
    const first = await api.listChatMessages("chat-1");
    expect(first.nextLink).toBe(nextLink);
    const second = await api.nextMessagePage(first.nextLink!);

    expect(second.items[0].id).toBe("1700000000001");
    expect(second.nextLink).toBeUndefined();
    // Graph put the ordering and its continuation token in the link; nothing is re-appended.
    expect(calls[1].url).toBe(nextLink);
  });

  it("refuses an off-origin nextLink returned by a listing", async () => {
    stubFetch(() => jsonResponse({
      value: [],
      "@odata.nextLink": "https://graph.microsoft.com.evil.example/v1.0/me/chats",
    }));

    await expect(newApi().listChats())
      .rejects.toThrow(/outside https:\/\/graph.microsoft.com/);
  });

  it("collects a listing that ends within its bounds", async () => {
    const nextLink = "https://graph.microsoft.com/v1.0/teams/team-1/channels?$skiptoken=next";
    const calls = stubFetch(call => call.url.includes("skiptoken")
      ? jsonResponse({ value: [{ id: "channel-2", displayName: "Random" }] })
      : jsonResponse({
        value: [{ id: "channel-1", displayName: "General", membershipType: "standard" }],
        "@odata.nextLink": nextLink,
      }));

    const channels = await newApi().listChannels("team-1");

    expect(channels.map(channel => channel.id)).toEqual(["channel-1", "channel-2"]);
    expect(calls).toHaveLength(2);
    // A channel does not carry its team, so the id comes from the request that found it.
    expect(channels[0].teamId).toBe("team-1");
  });

  it("refuses to answer with a listing truncated at the request cap", async () => {
    let page = 0;
    const calls = stubFetch(() => jsonResponse({
      value: [{ id: `team-${page++}`, displayName: "Team" }],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/joinedTeams?$skiptoken=next",
    }));

    // A short list of teams reads exactly like a complete one, so it is an error, not an answer.
    await expect(newApi().listJoinedTeams()).rejects.toThrow(/more joined teams than this can list/);
    expect(calls).toHaveLength(10);
  });

  it("refuses to answer with a listing truncated at the item cap", async () => {
    const page = Array.from({ length: 100 }, (_, index) => ({
      id: `channel-${index}`, displayName: "General", membershipType: "standard",
    }));
    const calls = stubFetch(() => jsonResponse({
      value: page,
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/teams/team-1/channels?$skiptoken=next",
    }));

    await expect(newApi().listChannels("team-1"))
      .rejects.toThrow(/more channels in this team than this can list/);
    // The item cap bites before the request cap does.
    expect(calls).toHaveLength(2);
  });
});

describe("error taxonomy", () => {
  it("truncates provider error text", async () => {
    stubFetch(() => jsonResponse({ error: { code: "invalidRequest", message: "x".repeat(600) } },
      400));

    const error = await newApi().listChats().catch(err => err) as GraphApiError;

    expect(error).toBeInstanceOf(GraphApiError);
    expect(error.message).toContain(`${"x".repeat(500)}…`);
    expect(error.message).not.toContain("x".repeat(501));
  });

  it("reports a claims-challenge 401 as credential death and does not retry", async () => {
    const rejected: string[] = [];
    const calls = stubFetch(() => jsonResponse(
      { error: { code: "InvalidAuthenticationToken", message: "policy" } }, 401, {
        "WWW-Authenticate":
          'Bearer realm="", error="insufficient_claims", claims="eyJhY2Nlc3NfdG9rZW4ifQ=="',
      }));

    const api = newApi({ onCredentialsRejected: async detail => { rejected.push(detail); } });
    const error = await api.listJoinedTeams().catch(err => err) as GraphApiError;

    expect(error.credentialsRejected).toBe(true);
    expect(error.message).toMatch(/sign in again/i);
    expect(rejected).toEqual(["insufficient_claims"]);
    expect(calls).toHaveLength(1);
  });

  it("points a 403 at the administrator consent the Teams scopes need", async () => {
    stubFetch(() => jsonResponse(
      { error: { code: "Forbidden", message: "Missing role permissions" } }, 403));

    const error = await newApi().listChannelMessages("team-1", "channel-1")
      .catch(err => err) as GraphApiError;

    expect(error.status).toBe(403);
    expect(error.message).toMatch(/administrator/i);
    expect(error.credentialsRejected).toBe(false);
  });

  it("explains a 404 in Teams terms", async () => {
    stubFetch(() => jsonResponse({ error: { code: "NotFound", message: "gone" } }, 404));

    const error = await newApi().getChat("chat-1").catch(err => err) as GraphApiError;

    expect(error.status).toBe(404);
    expect(error.message).toMatch(/no longer exists, or this account cannot see it/i);
  });
});

describe("throttling", () => {
  it("waits out a Retry-After the mailbox client would have clamped", async () => {
    vi.useFakeTimers();
    try {
      let attempts = 0;
      stubFetch(() => ++attempts === 1
        ? jsonResponse({ error: { code: "ApplicationThrottled" } }, 429, { "Retry-After": "30" })
        : jsonResponse({ value: [] }));

      const pending = newApi().listChats();
      await vi.advanceTimersByTimeAsync(10_000);
      // The shared 10s clamp would already have replayed here, straight back into the throttle.
      expect(attempts).toBe(1);

      await vi.advanceTimersByTimeAsync(20_000);
      await expect(pending).resolves.toMatchObject({ items: [] });
      expect(attempts).toBe(2);
    } finally {
      vi.useRealTimers();
    }
  });

  it("fails fast past the ceiling, saying how long to stay away", async () => {
    const calls = stubFetch(() => jsonResponse(
      { error: { code: "ApplicationThrottled" } }, 429, { "Retry-After": "180" }));

    const error = await newApi().listChats().catch(err => err) as GraphApiError;

    expect(error).toBeInstanceOf(GraphApiError);
    expect(error.status).toBe(429);
    expect(error.message).toMatch(/180 seconds/);
    expect(calls).toHaveLength(1);
  });

  it("leaves a 5xx Retry-After on the shared backoff instead of calling it throttling", async () => {
    vi.useFakeTimers();
    try {
      let attempts = 0;
      stubFetch(() => ++attempts === 1
        ? jsonResponse({ error: { code: "serviceError" } }, 503, { "Retry-After": "180" })
        : jsonResponse({ value: [] }));

      const pending = newApi().listChats();
      // A 5xx Retry-After is a hint about an outage, not a wait Graph is holding us to: it is
      // clamped and replayed rather than reported to the caller as a throttling ceiling.
      await vi.advanceTimersByTimeAsync(10_000);

      await expect(pending).resolves.toMatchObject({ items: [] });
      expect(attempts).toBe(2);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("search request body", () => {
  it("lets the query string in and nothing else", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [] }));
    const hostile = '"}],"requests":[{"entityTypes":["driveItem"],"query":{"queryString":"*"';

    await newApi().searchMessages(hostile, 7);

    expect(calls[0].url).toBe("https://graph.microsoft.com/v1.0/search/query");
    expect(calls[0].init.method).toBe("POST");
    expect(new Headers(calls[0].init.headers).get("Content-Type")).toBe("application/json");

    const body = JSON.parse(String(calls[0].init.body));
    expect(body).toEqual({
      requests: [{
        entityTypes: ["chatMessage"],
        query: { queryString: hostile },
        from: 7,
        size: 25,
      }],
    });
    // Blank out the one caller value and what remains is a literal skeleton.
    body.requests[0].query.queryString = "";
    expect(body).toEqual({
      requests: [{ entityTypes: ["chatMessage"], query: { queryString: "" }, from: 0 + 7, size: 25 }],
    });
  });

  it("clamps the offset it is given", () => {
    expect(JSON.parse(buildSearchRequestBody("q", -5)).requests[0].from).toBe(0);
    expect(JSON.parse(buildSearchRequestBody("q", 1e9)).requests[0].from).toBe(500);
    expect(JSON.parse(buildSearchRequestBody("q", Number.NaN)).requests[0].from).toBe(0);
    expect(JSON.parse(buildSearchRequestBody("q", 12.7)).requests[0].from).toBe(12);
  });
});

describe("search query validation", () => {
  it("allows a quoted phrase, which the mailbox validator cannot", () => {
    // The query travels in a JSON body, so a double quote is a KQL phrase, not an escape.
    expect(validateTeamsSearchQuery('"quarterly report" AND from:bob'))
      .toBe('"quarterly report" AND from:bob');
    expect(validateTeamsSearchQuery("  hasAttachment:true  ")).toBe("hasAttachment:true");
  });

  it("rejects what cannot be a real query", () => {
    expect(() => validateTeamsSearchQuery("   ")).toThrow(/must not be empty/);
    expect(() => validateTeamsSearchQuery("a".repeat(401))).toThrow(/at most 400/);
    expect(() => validateTeamsSearchQuery("from:bob\u0000drop")).toThrow(/control characters/);
    expect(() => validateTeamsSearchQuery("line\u001bescape")).toThrow(/control characters/);
  });

  it("validates before opening a walk, so a bad query costs no request", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [] }));

    expect(() => newApi().openMessageSearch("")).toThrow(/must not be empty/);

    expect(calls).toHaveLength(0);
  });
});

function searchResponse(
    hits: { id: string; channel?: [string, string]; chat?: string; summary?: string }[],
    moreResultsAvailable = false) {
  return {
    value: [{
      hitsContainers: [{
        total: hits.length,
        moreResultsAvailable,
        hits: hits.map((hit, index) => ({
          hitId: hit.id,
          rank: index + 1,
          summary: hit.summary ?? "the <c0>quarterly</c0> report",
          resource: {
            id: hit.id,
            createdDateTime: "2026-08-01T10:00:00Z",
            lastModifiedDateTime: "2026-08-02T10:00:00Z",
            from: { user: { id: "user-1", displayName: "Bob" } },
            subject: "Q3",
            webUrl: `https://teams.microsoft.com/l/message/${hit.id}`,
            ...(hit.channel ? { channelIdentity: { teamId: hit.channel[0], channelId: hit.channel[1] } } : {}),
            ...(hit.chat ? { chatId: hit.chat } : {}),
          },
        })),
      }],
    }],
  };
}

describe("search results", () => {
  it("maps a hit to metadata only, with the location the index reported", async () => {
    stubFetch(() => jsonResponse(searchResponse([
      { id: "m-1", channel: ["team-1", "channel-1"] },
      { id: "m-2", chat: "chat-1" },
      { id: "m-3" },
    ])));

    const page = await newApi().searchMessages("quarterly");

    expect(page.hits).toHaveLength(3);
    expect(page.hits[0]).toEqual({
      messageId: "m-1",
      from: { id: "user-1", displayName: "Bob", isApplication: false },
      createdAt: new Date("2026-08-01T10:00:00Z"),
      lastModifiedAt: new Date("2026-08-02T10:00:00Z"),
      // The index's highlight markup is stripped; the fragment itself is kept.
      summary: "the quarterly report",
      subject: "Q3",
      location: { kind: "channel", teamId: "team-1", channelId: "channel-1" },
      webUrl: "https://teams.microsoft.com/l/message/m-1",
    });
    expect(page.hits[1].location).toEqual({ kind: "chat", chatId: "chat-1" });
    // Reported, not dropped: the metadata is still worth showing even with nowhere to read it.
    expect(page.hits[2].location).toEqual({ kind: "unknown" });
  });

  it("drops a hit that names no message", async () => {
    stubFetch(() => jsonResponse({
      value: [{ hitsContainers: [{ hits: [{ hitId: "h-1", summary: "orphan" }] }] }],
    }));

    await expect(newApi().searchMessages("quarterly")).resolves.toEqual({
      hits: [], moreAvailable: false,
    });
  });

  it("pages by offset and never returns the same hit twice", async () => {
    const pages = [
      searchResponse([{ id: "m-1", chat: "chat-1" }, { id: "m-2", chat: "chat-1" }], true),
      // The corpus shifted under the offset: m-2 comes back on the next page.
      searchResponse([{ id: "m-2", chat: "chat-1" }, { id: "m-3", chat: "chat-1" }], false),
    ];
    let index = 0;
    const calls = stubFetch(() => jsonResponse(pages[index++]));

    const walk = newApi().openMessageSearch("quarterly");

    expect((await walk.peekPage())!.map(hit => hit.messageId)).toEqual(["m-1", "m-2"]);
    walk.commitPage();
    expect((await walk.peekPage())!.map(hit => hit.messageId)).toEqual(["m-3"]);
    walk.commitPage();
    expect(await walk.peekPage()).toBeNull();

    expect(JSON.parse(String(calls[0].init.body)).requests[0].from).toBe(0);
    expect(JSON.parse(String(calls[1].init.body)).requests[0].from).toBe(25);
    // Exhausted by moreResultsAvailable: no third request.
    expect(calls).toHaveLength(2);
  });

  it("leaves the walk where it was when a page is never committed", async () => {
    let index = 0;
    const pages = [
      searchResponse([{ id: "m-1", chat: "chat-1" }], true),
      searchResponse([{ id: "m-2", chat: "chat-1" }], true),
    ];
    const calls = stubFetch(() => jsonResponse(pages[Math.min(index++, 1)]));

    const walk = newApi().openMessageSearch("quarterly");

    expect((await walk.peekPage())!.map(hit => hit.messageId)).toEqual(["m-1"]);
    // The caller could not deliver that page — a refused approval, say — so it never committed.
    expect((await walk.peekPage())!.map(hit => hit.messageId)).toEqual(["m-1"]);
    // Re-answered from what the first peek already fetched, so the offset never moved either.
    expect(calls).toHaveLength(1);

    walk.commitPage();
    expect((await walk.peekPage())!.map(hit => hit.messageId)).toEqual(["m-2"]);
    expect(JSON.parse(String(calls[1].init.body)).requests[0].from).toBe(25);
  });

  it("answers overlapping peeks with one page rather than skipping an offset", async () => {
    let index = 0;
    const pages = [
      searchResponse([{ id: "m-1", chat: "chat-1" }], true),
      searchResponse([{ id: "m-2", chat: "chat-1" }], true),
    ];
    const calls = stubFetch(() => jsonResponse(pages[Math.min(index++, 1)]));

    const walk = newApi().openMessageSearch("quarterly");
    const [first, second] = await Promise.all([walk.peekPage(), walk.peekPage()]);

    // Unserialized, both would have requested offset 0 and one of them would have committed it,
    // leaving offset 25 never requested by anyone.
    expect(first!.map(hit => hit.messageId)).toEqual(["m-1"]);
    expect(second).toEqual(first);
    expect(calls).toHaveLength(1);

    walk.commitPage();
    expect((await walk.peekPage())!.map(hit => hit.messageId)).toEqual(["m-2"]);
    expect(JSON.parse(String(calls[1].init.body)).requests[0].from).toBe(25);
  });

  it("throws at the page cap rather than reporting the results as complete", async () => {
    let index = 0;
    const calls = stubFetch(() =>
      jsonResponse(searchResponse([{ id: `m-${index++}`, chat: "chat-1" }], true)));

    const walk = newApi().openMessageSearch("quarterly");
    for (let page = 0; page < 20; page++) {
      expect(await walk.peekPage()).not.toBeNull();
      walk.commitPage();
    }

    // Truncation is not exhaustion: null here would tell the agent it had seen everything.
    await expect(walk.peekPage()).rejects.toThrow(/Narrow the query/);
    expect(calls).toHaveLength(20);
  });

  it("reads the index's result-window ceiling as the end of the results", async () => {
    let index = 0;
    stubFetch(() => index++ === 0
      ? jsonResponse(searchResponse([{ id: "m-1", chat: "chat-1" }], true))
      : jsonResponse({
        error: {
          code: "invalidRequest",
          message: "Result window is too large, from + size must be less than or equal to 500.",
        },
      }, 400));

    const walk = newApi().openMessageSearch("quarterly");

    expect((await walk.peekPage())!.map(hit => hit.messageId)).toEqual(["m-1"]);
    walk.commitPage();
    // Not a failure: there is simply nothing past the window.
    expect(await walk.peekPage()).toBeNull();
  });

  it("reports a window-shaped failure on the first request as the error it is", async () => {
    stubFetch(() => jsonResponse({
      error: {
        code: "invalidRequest",
        message: "Result window is too large, from + size must be less than or equal to 500.",
      },
    }, 400));

    // Offset zero cannot have crossed a paging window, so this text is Graph rejecting the query.
    await expect(newApi().openMessageSearch("quarterly").peekPage())
      .rejects.toThrow(/Result window is too large/);
  });

  it("still reports a real search failure", async () => {
    stubFetch(() => jsonResponse(
      { error: { code: "invalidRequest", message: "The query is malformed." } }, 400));

    await expect(newApi().openMessageSearch("quarterly").peekPage())
      .rejects.toThrow(/malformed/);
  });
});

describe("search hit hydration", () => {
  function hit(location: TeamsSearchHitInfo["location"]): TeamsSearchHitInfo {
    return {
      messageId: "1700000000000",
      createdAt: new Date("2026-08-01T10:00:00Z"),
      summary: "the quarterly report",
      location,
    };
  }

  it("reads a hit from the conversation the index named", async () => {
    const calls = stubFetch(() => jsonResponse(CHANNEL_MESSAGE));
    const api = newApi();

    await api.getSearchHitMessage(hit({ kind: "channel", teamId: "t-1", channelId: "c-1" }));
    await api.getSearchHitMessage(hit({ kind: "chat", chatId: "chat-1" }));

    expect(new URL(calls[0].url).pathname)
      .toBe("/v1.0/teams/t-1/channels/c-1/messages/1700000000000");
    expect(new URL(calls[1].url).pathname).toBe("/v1.0/chats/chat-1/messages/1700000000000");
  });

  it("fails loudly for a hit with no resolvable location", async () => {
    const calls = stubFetch(() => jsonResponse(CHANNEL_MESSAGE));

    const error = await newApi().getSearchHitMessage(hit({ kind: "unknown" }))
      .catch(err => err) as GraphApiError;

    expect(error).toBeInstanceOf(GraphApiError);
    expect(error.message).toMatch(/location could not be resolved/i);
    expect(calls).toHaveLength(0);
  });
});

describe("body rendering", () => {
  it("resolves a mention and an attachment against the message's own lists", () => {
    const body = renderTeamsMessageBody({
      body: html('<div><at id="0">Jane Doe</at> please review ' +
        '<attachment id="a1"></attachment></div>'),
      mentions: [{
        id: 0, mentionText: "Jane Doe", mentioned: { user: { id: "u-1", displayName: "Jane Doe" } },
      }],
      attachments: [{ id: "a1", name: "Q3.docx", contentType: "reference" }],
    });

    expect(body).toBe("@Jane Doe please review [attachment: Q3.docx]");
  });

  it("falls back to what the body itself said when a mention is not listed", () => {
    expect(renderTeamsMessageBody({ body: html('hi <at id="7">General</at>') }))
      .toBe("hi @General");
  });

  it("names an attachment by content type, or not at all", () => {
    expect(renderTeamsMessageBody({
      body: html('<attachment id="a1"></attachment>'),
      attachments: [{ id: "a1", contentType: "application/vnd.microsoft.card.adaptive" }],
    })).toBe("[attachment: application/vnd.microsoft.card.adaptive]");

    expect(renderTeamsMessageBody({ body: html('<attachment id="missing"></attachment>') }))
      .toBe("[attachment]");
  });

  it("renders an emoji as the alt text Teams carries", () => {
    expect(renderTeamsMessageBody({
      body: html('<p>Nice <emoji id="smile" alt="\u{1F600}" title="Grinning"></emoji></p>'),
    })).toBe("Nice \u{1F600}");
    expect(renderTeamsMessageBody({ body: html('<emoji id="x" title="Grinning"></emoji>') }))
      .toBe("Grinning");
  });

  it("renders a system event's marker as nothing", () => {
    expect(renderTeamsMessageBody({ body: html("<systemEventMessage/>") })).toBe("");
  });

  it("strips a tag it does not know and keeps the text around it", () => {
    expect(renderTeamsMessageBody({
      body: html('<div>before <weird data-x="1">inside</weird> after</div>'),
    })).toBe("before inside after");
  });

  it("decodes entities, leaving ones it does not know as written", () => {
    expect(renderTeamsMessageBody({
      body: html("a &amp; b &lt;tag&gt; &#233; &#x1F600; &unknownthing;"),
    })).toBe("a & b <tag> é \u{1F600} &unknownthing;");
  });

  it("keeps block markup as line breaks", () => {
    expect(renderTeamsMessageBody({ body: html("<p>one</p><p>two</p><div>three<br>four</div>") }))
      .toBe("one\n\ntwo\n\nthree\nfour");
  });

  it("passes a plain-text body through untouched", () => {
    expect(renderTeamsMessageBody({
      body: { contentType: "text", content: "  a < b &amp; c  " },
    })).toBe("a < b &amp; c");
  });

  it("caps a pathological body", () => {
    const body = renderTeamsMessageBody({ body: html("a".repeat(40 * 1024)) });

    expect(body).toHaveLength(32 * 1024 + 1);
    expect(body.endsWith("…")).toBe(true);
  });

  it("never throws on malformed markup", () => {
    for (const content of [
      "<div>unclosed", '<at id="0">Jane', "<<>>", "<!-- comment", "a < b", "</at></div>",
      '<attachment id="a1"', "<emoji alt=\u{1F600}>",
    ]) {
      expect(() => renderTeamsMessageBody({ body: html(content) })).not.toThrow();
    }
  });

  it("keeps the body that an unclosed mention would otherwise swallow", () => {
    // `<at>` that never closes was never a mention: what it collected is the message text.
    expect(renderTeamsMessageBody({ body: html('<at id="0">Jane') })).toBe("Jane");
    expect(renderTeamsMessageBody({
      body: html('<div>hi <at id="0">Jane<p>the rest of what was said &amp; meant</p>'),
    })).toBe("hi Jane\nthe rest of what was said & meant");
  });

  it("bounds a mention label instead of diverting the body into one", () => {
    const body = renderTeamsMessageBody({
      body: html(`<at id="0">${"n".repeat(200)} and then the message itself`),
    });

    // The label stops at its ceiling and the text past it is still text.
    expect(body).toContain("and then the message itself");
    expect(body.length).toBeGreaterThan(200);
  });

  it("keeps a hostile attachment name from reshaping the body", () => {
    const body = renderTeamsMessageBody({
      body: html('<attachment id="a1"></attachment>'),
      attachments: [{ id: "a1", name: `evil\n\n${"n".repeat(200)}` }],
    });

    expect(body).not.toContain("\n");
    expect(body.length).toBeLessThan(160);
  });
});

describe("message mapping", () => {
  it("maps a channel message onto the agent-facing shape", async () => {
    stubFetch(() => jsonResponse({ value: [CHANNEL_MESSAGE] }));

    const page = await newApi().listChannelMessages("team-1", "channel-1");

    expect(page.items[0]).toEqual({
      id: "1700000000000",
      from: { id: "user-1", displayName: "Bob", isApplication: false },
      createdAt: new Date("2026-08-01T10:00:00Z"),
      lastEditedAt: new Date("2026-08-01T10:05:00Z"),
      messageKind: "message",
      body: "Numbers attached",
      attachments: [],
      mentions: [],
      deleted: false,
      webUrl: "https://teams.microsoft.com/l/message/19:channel/1700000000000",
    });
  });

  it("reports a deleted message with an empty body, and a system notice as one", async () => {
    stubFetch(() => jsonResponse({
      value: [
        {
          ...CHANNEL_MESSAGE,
          id: "deleted-1",
          deletedDateTime: "2026-08-03T10:00:00Z",
          body: html("<div>was here</div>"),
        },
        {
          id: "system-1",
          // Without `Prefer: include-unknown-enum-members`, which this client does not send, Graph
          // reports a system event's type as this: `eventDetail` is what identifies it.
          messageType: "unknownFutureValue",
          createdDateTime: "2026-08-01T09:00:00Z",
          body: html("<systemEventMessage/>"),
          eventDetail: { "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail" },
        },
      ],
    }));

    const page = await newApi().listChannelMessages("team-1", "channel-1");

    expect(page.items[0]).toMatchObject({ deleted: true, body: "", messageKind: "message" });
    expect(page.items[1]).toMatchObject({
      deleted: false, messageKind: "systemEvent", body: "[system event: members added]",
    });
    // A system notice has no author, so the field is absent rather than a placeholder person.
    expect(page.items[1]).not.toHaveProperty("from");
  });

  it("marks a bot post as coming from an application", async () => {
    stubFetch(() => jsonResponse({
      value: [{ ...CHANNEL_MESSAGE, from: BOT_IDENTITY }],
    }));

    const page = await newApi().listChatMessages("chat-1");

    // Graph sends the unused identities as explicit nulls, not as missing keys.
    expect(page.items[0].from).toEqual({
      id: "app-1", displayName: "Builds", isApplication: true,
    });
  });

  it("maps members and chats onto their agent-facing shapes", async () => {
    stubFetch(call => new URL(call.url).pathname.endsWith("/members")
      ? jsonResponse({ value: [{
        id: "membership-1", userId: "user-1", displayName: "Bob", email: "bob@example.com",
        roles: ["owner"],
      }] })
      : jsonResponse({ value: [{
        id: "chat-1", chatType: "meeting", topic: "Standup",
        createdDateTime: "2026-07-01T10:00:00Z", lastUpdatedDateTime: "2026-08-01T10:00:00Z",
      }] }));

    const api = newApi();

    expect((await api.listChatMembers("chat-1")).items[0]).toEqual({
      id: "membership-1", userId: "user-1", displayName: "Bob", email: "bob@example.com",
      roles: ["owner"],
    });
    expect((await api.listChats()).items[0]).toEqual({
      id: "chat-1",
      chatKind: "meeting",
      topic: "Standup",
      createdAt: new Date("2026-07-01T10:00:00Z"),
      lastUpdatedAt: new Date("2026-08-01T10:00:00Z"),
    });
  });

  it("refuses a payload with no id rather than inventing one", async () => {
    stubFetch(() => jsonResponse({ value: [{ displayName: "Nameless" }] }));

    await expect(newApi().listJoinedTeams()).rejects.toThrow(/without an id/);
  });
});

describe("system event notices", () => {
  const SYSTEM_MESSAGE: GraphChatMessage = {
    id: "system-1",
    messageType: "unknownFutureValue",
    createdDateTime: "2026-08-01T09:00:00Z",
    body: html("<systemEventMessage/>"),
  };

  async function read(message: GraphChatMessage) {
    stubFetch(() => jsonResponse(message));
    return await newApi().getChatMessage("chat-1", "system-1");
  }

  it("names the event, derived from the detail type Graph sends", async () => {
    expect((await read({
      ...SYSTEM_MESSAGE,
      eventDetail: { "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail" },
    })).body).toBe("[system event: members added]");

    // Derived, not looked up: an event type this client has never heard of still reads as itself.
    const favorited = "#microsoft.graph.channelSetAsFavoriteByDefaultEventMessageDetail";
    expect((await read({ ...SYSTEM_MESSAGE, eventDetail: { "@odata.type": favorited } })).body)
      .toBe("[system event: channel set as favorite by default]");
  });

  it("says only that something happened when the detail is not an event type", async () => {
    const message = await read({
      ...SYSTEM_MESSAGE,
      eventDetail: { "@odata.type": "#microsoft.graph.somethingElseEntirely" },
    });

    expect(message.body).toBe("[system event]");
    expect(message.messageKind).toBe("systemEvent");
  });

  it("says only that something happened when Graph sends no detail at all", async () => {
    const message = await read({ ...SYSTEM_MESSAGE, messageType: "systemEventMessage" });

    // The type alone still marks it a system event; there is just nothing to name.
    expect(message.messageKind).toBe("systemEvent");
    expect(message.body).toBe("[system event]");
  });

  it("leaves a deleted system message empty rather than describing it", async () => {
    const message = await read({
      ...SYSTEM_MESSAGE,
      deletedDateTime: "2026-08-03T10:00:00Z",
      eventDetail: { "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail" },
    });

    expect(message).toMatchObject({ deleted: true, messageKind: "systemEvent", body: "" });
  });
});
