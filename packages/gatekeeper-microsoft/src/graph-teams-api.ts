// Microsoft Graph Teams client.
//
// Same shape as the mailbox client next door (`graph-api.ts`): one request chokepoint, URLs built
// only by `graphUrl`, pagination links pinned to Graph's origin, bounded pages, truncated provider
// text. What differs is Teams-specific and deliberate:
//
//   - No `Prefer: IdType="ImmutableId"`, and no body-content-type preference. Both are Outlook
//     options: mail ids change when a message moves folders, and Outlook can render a body as text
//     on request. Teams ids are stable, and Teams answers with its own pseudo-HTML no matter what is
//     asked for, so the rendering happens here instead (`renderTeamsMessageBody`).
//   - A real `Retry-After` ceiling. Teams throttling routinely asks for a minute or more, while the
//     shared backoff clamps every wait to ten seconds — which would replay straight back into the
//     throttle. This client waits out up to `MAX_RETRY_AFTER_MS` and beyond that fails fast, saying
//     how long to stay away.
//   - Search is a POST, so the injection surface is a JSON body rather than a URL. Every search body
//     is written by `buildSearchRequestBody`, a fixed skeleton with exactly one caller-supplied
//     value in it, and the query itself is validated by a Teams-local validator: the mailbox's
//     validator bans double quotes to protect a quoted `$search` URL parameter, which would only
//     break KQL phrase search here.
//
// Reads only. Nothing in this file posts, edits, deletes, or marks anything.

import {
  AccessTokenProvider, CredentialsRejectedReporter, RetryAfterPolicy, claimsChallengeDetail,
  fetchWithAuthRetry,
} from "./auth-retry";
import {
  GRAPH_TIMEOUT_MS, GraphApiError, GraphCollection, GraphPage, assertGraphUrl, graphUrl, truncate,
} from "./graph-api";
import type {
  TeamsAttachmentInfo, TeamsChannelInfo, TeamsChannelMembership, TeamsChatInfo, TeamsChatKind,
  TeamsMemberInfo, TeamsMentionInfo, TeamsMessageInfo, TeamsMessageKind, TeamsSearchHitInfo,
  TeamsSearchHitLocation, TeamsTeamInfo, TeamsUser,
} from "./teams-types";

/** Messages Graph hands out per page. The `$top` ceiling on Teams collections is 50. */
const TEAMS_MESSAGE_PAGE_SIZE = 25;
const CHATS_PAGE_SIZE = 50;
const MEMBERS_PAGE_SIZE = 50;
const MAX_PAGE_SIZE = 50;

/**
 * Ceilings on the listings that page internally rather than through a cursor (teams and channels,
 * which the session API returns as whole arrays). Bounded by items and by requests.
 */
const MAX_TEAMS = 200;
const MAX_CHANNELS = 200;
const MAX_WALK_PAGES = 10;

/** Longest KQL search string accepted from an agent. */
const MAX_SEARCH_QUERY_CHARS = 400;

/** Search paging is by offset, so both the window and the walk over it are bounded here. */
const SEARCH_PAGE_SIZE = 25;
const MAX_SEARCH_PAGES = 20;
const MAX_SEARCH_OFFSET = MAX_SEARCH_PAGES * SEARCH_PAGE_SIZE;

/**
 * Longest throttling wait this client sits through. Beyond it the request fails with the wait in the
 * message: an agent that is told "come back in four minutes" can do something else, while a client
 * that quietly slept would look hung and then get throttled again anyway.
 */
const MAX_RETRY_AFTER_MS = 60 * 1000;

/** Longest rendered message body handed to an agent. */
const MAX_MESSAGE_BODY_CHARS = 32 * 1024;

/** Longest counterpart-controlled label (attachment name, mention text) placed inline in a body. */
const MAX_INLINE_LABEL_CHARS = 120;

// ── Graph wire shapes ───────────────────────────────────────────────
// Only the fields this client reads. Everything is optional: these describe what Graph may send, not
// what it promises.

type GraphIdentity = {
  id?: string;
  displayName?: string;
  email?: string;
  userIdentityType?: string;
};

/**
 * An identity set names at most one of a user, an application, and a device. Graph sends the ones
 * that do not apply as explicit `null`s rather than leaving them out, so every reader here has to
 * treat null and absent alike.
 */
type GraphIdentitySet = {
  user?: GraphIdentity | null;
  application?: GraphIdentity | null;
  device?: GraphIdentity | null;
};

type GraphTeam = {
  id?: string;
  displayName?: string;
  description?: string;
  webUrl?: string;
};

type GraphChannel = {
  id?: string;
  displayName?: string;
  description?: string;
  membershipType?: string;
  webUrl?: string;
};

type GraphChat = {
  id?: string;
  topic?: string;
  chatType?: string;
  createdDateTime?: string;
  lastUpdatedDateTime?: string;
  webUrl?: string;
};

type GraphConversationMember = {
  id?: string;
  displayName?: string;
  roles?: string[];
  userId?: string;
  email?: string;
};

type GraphMessageAttachment = {
  id?: string;
  contentType?: string;
  contentUrl?: string;
  name?: string;
};

type GraphMessageMention = {
  id?: number;
  mentionText?: string;
  mentioned?: GraphIdentitySet;
};

/** A `chatMessage`, in a channel or a chat. Search returns a metadata-only subset of the same type. */
export type GraphChatMessage = {
  id?: string;
  messageType?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  lastEditedDateTime?: string;
  deletedDateTime?: string;
  subject?: string;
  summary?: string;
  webUrl?: string;
  from?: GraphIdentitySet;
  body?: { contentType?: string; content?: string };
  attachments?: GraphMessageAttachment[];
  mentions?: GraphMessageMention[];
  eventDetail?: Record<string, unknown>;
  /** Present on channel messages; identifies the conversation a search hit came from. */
  channelIdentity?: { teamId?: string; channelId?: string };
  /** Present on chat messages, for the same reason. */
  chatId?: string;
};

type GraphSearchHit = {
  hitId?: string;
  rank?: number;
  summary?: string;
  resource?: GraphChatMessage;
};

type GraphSearchResponse = {
  value?: {
    hitsContainers?: {
      hits?: GraphSearchHit[];
      total?: number;
      moreResultsAvailable?: boolean;
    }[];
  }[];
};

// ── Mapping onto the agent-facing types ─────────────────────────────

function text(value: unknown): string | undefined {
  return typeof value === "string" && value ? value : undefined;
}

function dateFrom(value: unknown): Date | undefined {
  let raw = text(value);
  if (!raw) return undefined;
  let parsed = new Date(raw);
  return Number.isNaN(parsed.valueOf()) ? undefined : parsed;
}

/**
 * The person or app behind an identity set.
 *
 * Teams reports a sender as a set of identities of which at most one is filled in. A message with no
 * user identity but an application one came from a bot or connector, which is worth telling an agent
 * apart from a colleague — and "no user identity" means null as often as absent, which is why the
 * test below is falsiness and not `=== undefined`.
 */
function userFrom(identity: GraphIdentitySet | undefined): TeamsUser | undefined {
  if (!identity) return undefined;
  let source = identity.user ?? identity.application ?? identity.device;
  if (!source) return undefined;
  let id = text(source.id);
  let email = text(source.email);
  return {
    ...(id ? { id } : {}),
    displayName: text(source.displayName) ?? "(unknown)",
    ...(email ? { email } : {}),
    isApplication: !identity.user,
  };
}

function teamInfoFrom(team: GraphTeam): TeamsTeamInfo {
  let id = text(team.id);
  if (!id) throw new Error("Microsoft Graph returned a team without an id.");
  let description = text(team.description);
  let webUrl = text(team.webUrl);
  return {
    id,
    displayName: text(team.displayName) ?? "(unnamed team)",
    ...(description ? { description } : {}),
    ...(webUrl ? { webUrl } : {}),
  };
}

function membershipFrom(value: string | undefined): TeamsChannelMembership {
  switch (value) {
    case "standard": return "standard";
    case "private": return "private";
    case "shared": return "shared";
    default: return "unknown";
  }
}

// `teamId` comes from the request, not the payload: a channel resource does not carry its team.
function channelInfoFrom(teamId: string, channel: GraphChannel): TeamsChannelInfo {
  let id = text(channel.id);
  if (!id) throw new Error("Microsoft Graph returned a channel without an id.");
  let description = text(channel.description);
  let webUrl = text(channel.webUrl);
  return {
    id,
    teamId,
    displayName: text(channel.displayName) ?? "(unnamed channel)",
    ...(description ? { description } : {}),
    membershipType: membershipFrom(text(channel.membershipType)),
    ...(webUrl ? { webUrl } : {}),
  };
}

function chatKindFrom(value: string | undefined): TeamsChatKind {
  switch (value) {
    case "oneOnOne": return "oneOnOne";
    case "group": return "group";
    case "meeting": return "meeting";
    default: return "unknown";
  }
}

function chatInfoFrom(chat: GraphChat): TeamsChatInfo {
  let id = text(chat.id);
  if (!id) throw new Error("Microsoft Graph returned a chat without an id.");
  let topic = text(chat.topic);
  let createdAt = dateFrom(chat.createdDateTime);
  let lastUpdatedAt = dateFrom(chat.lastUpdatedDateTime);
  let webUrl = text(chat.webUrl);
  return {
    id,
    chatKind: chatKindFrom(text(chat.chatType)),
    ...(topic ? { topic } : {}),
    ...(createdAt ? { createdAt } : {}),
    ...(lastUpdatedAt ? { lastUpdatedAt } : {}),
    ...(webUrl ? { webUrl } : {}),
  };
}

function memberInfoFrom(member: GraphConversationMember): TeamsMemberInfo {
  let id = text(member.id);
  if (!id) throw new Error("Microsoft Graph returned a conversation member without an id.");
  let userId = text(member.userId);
  let email = text(member.email);
  return {
    id,
    ...(userId ? { userId } : {}),
    displayName: text(member.displayName) ?? "(unknown member)",
    ...(email ? { email } : {}),
    roles: Array.isArray(member.roles)
        ? member.roles.filter((role): role is string => typeof role === "string")
        : [],
  };
}

function attachmentsFrom(attachments: GraphMessageAttachment[] | undefined): TeamsAttachmentInfo[] {
  if (!Array.isArray(attachments)) return [];
  return attachments.flatMap(attachment => {
    let id = text(attachment.id);
    // An attachment with no id cannot be referred to by the body's `<attachment>` tag either, so
    // there is nothing an agent could do with it.
    if (!id) return [];
    let name = text(attachment.name);
    let contentType = text(attachment.contentType);
    let contentUrl = text(attachment.contentUrl);
    return [{
      id,
      ...(name ? { name } : {}),
      ...(contentType ? { contentType } : {}),
      ...(contentUrl ? { contentUrl } : {}),
    }];
  });
}

function mentionsFrom(mentions: GraphMessageMention[] | undefined): TeamsMentionInfo[] {
  if (!Array.isArray(mentions)) return [];
  return mentions.map(mention => {
    // Only a user mention resolves to a person; channel, team, and app mentions carry no user.
    let user = mention.mentioned?.user ? userFrom({ user: mention.mentioned.user }) : undefined;
    return {
      text: text(mention.mentionText) ?? user?.displayName ?? "(mention)",
      ...(user ? { user } : {}),
    };
  });
}

/**
 * What a message is.
 *
 * `eventDetail` is what decides it, before `messageType` is consulted at all: Graph fills it in for
 * system events only, and it does so whatever the type says. The type on its own is not enough,
 * because a client that does not ask for unknown enum members — this one does not — is answered
 * `unknownFutureValue` for every event kind Teams added after the API version, so a join or a rename
 * would be reported as "other".
 */
function messageKindFrom(message: GraphChatMessage): TeamsMessageKind {
  if (message.eventDetail) return "systemEvent";
  switch (text(message.messageType)) {
    case "systemEventMessage":
    case "chatEvent":
      return "systemEvent";
    case "message":
    case undefined:
      return "message";
    default: return "other";
  }
}

/**
 * A one-line notice saying what a system event was.
 *
 * A system message has no text of its own: the body is a marker tag and what happened is in
 * `eventDetail`, whose `@odata.type` names it — `#microsoft.graph.membersAddedEventMessageDetail`
 * becomes `[system event: members added]`. The words are derived from that name rather than looked
 * up in a table of the events Teams can raise today, so an event added tomorrow still reads as
 * itself. A name that is not one of those resource types says only that something happened, which is
 * all it is honest to say.
 *
 * The notice is the whole body of a system message, so it bypasses `normalizeBody`; the words are
 * bounded here instead, by the same `inlineLabel` every other provider-derived label goes through.
 */
function renderSystemEventNotice(message: GraphChatMessage): string {
  let odataType = text(message.eventDetail?.["@odata.type"]);
  let name = odataType?.match(/^#?microsoft\.graph\.(.+)EventMessageDetail$/i)?.[1] ?? "";
  let words = inlineLabel(name
      .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
      .toLowerCase()
      .replace(/[^a-z]+/g, " "));
  return words ? `[system event: ${words}]` : "[system event]";
}

function messageInfoFrom(message: GraphChatMessage): TeamsMessageInfo {
  let id = text(message.id);
  if (!id) throw new Error("Microsoft Graph returned a message without an id.");
  let from = userFrom(message.from);
  let lastEditedAt = dateFrom(message.lastEditedDateTime);
  let deleted = dateFrom(message.deletedDateTime) !== undefined;
  let kind = messageKindFrom(message);
  let webUrl = text(message.webUrl);
  return {
    id,
    ...(from ? { from } : {}),
    createdAt: dateFrom(message.createdDateTime) ?? new Date(0),
    ...(lastEditedAt ? { lastEditedAt } : {}),
    messageKind: kind,
    // A deleted message keeps its place in a listing, but Graph returns no body for it. A system
    // notice's body is only the marker tag, so the notice is synthesized from `eventDetail` — an
    // agent reading a listing would otherwise see a message with no text and no reason for it.
    body: deleted
        ? ""
        : kind === "systemEvent" ? renderSystemEventNotice(message) : renderTeamsMessageBody(message),
    attachments: attachmentsFrom(message.attachments),
    mentions: mentionsFrom(message.mentions),
    deleted,
    ...(webUrl ? { webUrl } : {}),
  };
}

/**
 * Where search says a hit lives.
 *
 * A hit whose index entry carries neither a channel identity nor a chat id is reported as `unknown`
 * rather than dropped: the metadata search did return is still worth showing, and hydrating such a
 * hit fails loudly (see `getSearchHitMessage`) instead of silently going missing from the results.
 */
function hitLocationFrom(resource: GraphChatMessage): TeamsSearchHitLocation {
  let teamId = text(resource.channelIdentity?.teamId);
  let channelId = text(resource.channelIdentity?.channelId);
  if (teamId && channelId) return { kind: "channel", teamId, channelId };
  let chatId = text(resource.chatId);
  if (chatId) return { kind: "chat", chatId };
  return { kind: "unknown" };
}

function searchHitInfoFrom(hit: GraphSearchHit): TeamsSearchHitInfo | undefined {
  let resource = hit.resource;
  let messageId = text(resource?.id);
  // Without an id the hit names no message, so there is nothing to report or hydrate.
  if (!resource || !messageId) return undefined;
  let from = userFrom(resource.from);
  let lastModifiedAt = dateFrom(resource.lastModifiedDateTime);
  let subject = text(resource.subject);
  let webUrl = text(resource.webUrl);
  return {
    messageId,
    ...(from ? { from } : {}),
    createdAt: dateFrom(resource.createdDateTime) ?? new Date(0),
    ...(lastModifiedAt ? { lastModifiedAt } : {}),
    // Summaries arrive with the index's own highlight markup around the matched words.
    summary: renderTeamsText(hit.summary ?? ""),
    ...(subject ? { subject } : {}),
    location: hitLocationFrom(resource),
    ...(webUrl ? { webUrl } : {}),
  };
}

// ── Body rendering ──────────────────────────────────────────────────
//
// Teams message bodies are pseudo-HTML: ordinary markup plus tags only Teams understands — `<at>`
// for mentions, `<attachment>` for files and cards, `<emoji>`, and `<systemEventMessage/>`. An agent
// needs plain text, and a Worker has no HTML parser, so this is a small scanner: text is
// entity-decoded, the four Teams tags are rendered from the message's own `mentions`/`attachments`,
// tags that mark a break become newlines, and everything else is dropped. Nothing here throws — an
// unrecognized or malformed tag degrades to "strip it and keep the text around it", because a body
// that cannot be rendered is still a body the agent should see.

const NAMED_ENTITIES: Record<string, string> = {
  amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'", nbsp: "\u00a0",
};

/** Tags whose presence means a line ended, whether they open or close. */
const BREAK_TAGS = new Set([
  "br", "p", "div", "li", "ul", "ol", "tr", "table", "blockquote", "pre",
  "h1", "h2", "h3", "h4", "h5", "h6",
]);

function codePointOr(code: number, fallback: string): string {
  // Lone surrogates and NUL are not text; leaving the entity as written is more honest than
  // emitting an unpaired code unit.
  if (!Number.isFinite(code) || code <= 0 || code > 0x10ffff) return fallback;
  if (code >= 0xd800 && code <= 0xdfff) return fallback;
  return String.fromCodePoint(code);
}

function decodeEntities(value: string): string {
  if (!value.includes("&")) return value;
  return value.replace(/&(#x[0-9a-f]+|#[0-9]+|[a-z][a-z0-9]*);/gi, (match, name: string) => {
    let lower = name.toLowerCase();
    if (lower.startsWith("#x")) return codePointOr(parseInt(lower.slice(2), 16), match);
    if (lower.startsWith("#")) return codePointOr(parseInt(lower.slice(1), 10), match);
    // An entity we do not know stays exactly as the author wrote it.
    return NAMED_ENTITIES[lower] ?? match;
  });
}

type TeamsTag = { name: string; closing: boolean; attrs: Record<string, string> };

/**
 * Index of the `>` that closes the tag opening at `start`, or -1 when the markup never closes it.
 * Quoted attribute values are skipped, so a `>` inside an attribute does not end the tag early.
 */
function tagEnd(html: string, start: number): number {
  let quote = "";
  for (let i = start + 1; i < html.length; i++) {
    let ch = html[i];
    if (quote) {
      if (ch === quote) quote = "";
      continue;
    }
    if (ch === "\"" || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === ">") return i;
  }
  return -1;
}

function parseTag(raw: string): TeamsTag {
  let closing = raw.startsWith("/");
  let body = closing ? raw.slice(1) : raw;
  let name = body.match(/^[a-z][a-z0-9:_-]*/i)?.[0].toLowerCase() ?? "";
  let attrs: Record<string, string> = {};
  let pattern = /([a-z_:][a-z0-9_:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+))/gi;
  for (let match of body.slice(name.length).matchAll(pattern)) {
    attrs[match[1].toLowerCase()] = decodeEntities(match[2] ?? match[3] ?? match[4] ?? "");
  }
  return { name, closing, attrs };
}

/** Flatten a counterpart-controlled label so it cannot reshape the rendered body. */
function inlineLabel(value: string): string {
  let flat = value.replace(/\s+/g, " ").trim();
  return flat.length > MAX_INLINE_LABEL_CHARS
      ? `${flat.slice(0, MAX_INLINE_LABEL_CHARS)}…`
      : flat;
}

function renderMention(
    id: string | undefined, innerText: string, message: GraphChatMessage): string {
  let mention = (message.mentions ?? []).find(entry => String(entry.id) === id);
  let name = text(mention?.mentionText)
      ?? text(mention?.mentioned?.user?.displayName)
      ?? text(innerText.trim());
  return name ? `@${inlineLabel(name)}` : "";
}

function renderAttachment(id: string | undefined, message: GraphChatMessage): string {
  let attachment = (message.attachments ?? []).find(entry => entry.id === id);
  let label = text(attachment?.name) ?? text(attachment?.contentType);
  return label ? `[attachment: ${inlineLabel(label)}]` : "[attachment]";
}

/** Collapse whitespace runs, drop the non-breaking spaces Teams sprinkles in, and cap the result. */
function normalizeBody(value: string): string {
  let collapsed = value
      .replace(/\r\n?/g, "\n")
      .replace(/\u00a0/g, " ")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  return collapsed.length > MAX_MESSAGE_BODY_CHARS
      ? `${collapsed.slice(0, MAX_MESSAGE_BODY_CHARS)}…`
      : collapsed;
}

/**
 * Render Teams pseudo-HTML to plain text. `message` supplies the `mentions` and `attachments` the
 * body's tags refer to by id; pass just the content (as `renderTeamsText` does) when there are none.
 */
function renderMarkup(content: string, message: GraphChatMessage): string {
  let out: string[] = [];
  let rendered = 0;
  // `<at>` collects its own inner text, so a mention Graph did not list still shows the name the
  // body carried. What it collects is a label, so it is bounded: an `<at>` that never closes must
  // not divert the rest of the body into a name.
  let mentionDepth = 0;
  let mentionText = "";
  let mentionId: string | undefined;

  let emit = (value: string) => {
    if (!value) return;
    let rest = value;
    if (mentionDepth > 0) {
      let room = MAX_INLINE_LABEL_CHARS - mentionText.length;
      mentionText += value.slice(0, room);
      if (value.length < room) return;
      // The label is full, so whatever this is, it is no longer a name being written: the mention
      // is rendered from what it has and the rest of the text goes on being text.
      endMention(true);
      rest = value.slice(room);
      if (!rest) return;
    }
    out.push(rest);
    rendered += rest.length;
  };
  let emitText = (value: string) => emit(decodeEntities(value));

  /**
   * Finish the open mention. As a mention it renders to `@name`; otherwise what it collected is
   * emitted as the ordinary text it turned out to be — which is what an `<at>` that never closed
   * was written around.
   */
  let endMention = (asMention: boolean) => {
    mentionDepth = 0;
    let inner = mentionText;
    let id = mentionId;
    mentionText = "";
    mentionId = undefined;
    emit(asMention ? renderMention(id, inner, message) : inner);
  };

  let index = 0;
  // Both bounds do move: `index` walks the markup below, and `rendered` grows inside `emit()`,
  // which every branch of this loop reaches.
  // oxlint-disable-next-line no-unmodified-loop-condition -- `emit()` advances `rendered`.
  while (index < content.length && rendered <= MAX_MESSAGE_BODY_CHARS) {
    let open = content.indexOf("<", index);
    if (open < 0) {
      emitText(content.slice(index));
      break;
    }
    if (open > index) emitText(content.slice(index, open));

    if (content.startsWith("<!--", open)) {
      let close = content.indexOf("-->", open);
      index = close < 0 ? content.length : close + 3;
      continue;
    }
    let end = tagEnd(content, open);
    if (end < 0) {
      // Never closed, so it is not a tag: show it as the text it is.
      emitText(content.slice(open));
      break;
    }
    let tag = parseTag(content.slice(open + 1, end));
    index = end + 1;

    switch (tag.name) {
      case "at":
        if (tag.closing) {
          if (mentionDepth > 0) endMention(true);
        } else {
          mentionDepth++;
          if (mentionDepth === 1) {
            mentionText = "";
            mentionId = tag.attrs["id"];
          }
        }
        break;
      case "attachment":
        if (!tag.closing) emit(renderAttachment(tag.attrs["id"], message));
        break;
      case "emoji":
        if (!tag.closing) emit(decodeEntities(tag.attrs["alt"] ?? tag.attrs["title"] ?? ""));
        break;
      case "systemeventmessage":
        // The marker carries no text of its own; what happened is in the message's `eventDetail`,
        // which `messageInfoFrom` renders into the body an agent reads.
        break;
      default:
        if (BREAK_TAGS.has(tag.name)) emit("\n");
        break;
    }
  }

  // A mention still open at the end of the markup never was one.
  if (mentionDepth > 0) endMention(false);
  return out.join("");
}

/** Render a Teams markup fragment that refers to no mentions or attachments, e.g. a search summary. */
function renderTeamsText(content: string): string {
  return normalizeBody(renderMarkup(content, {}));
}

/**
 * The message body as plain text: entities decoded, Teams' own tags resolved against the message's
 * mentions and attachments, all other markup stripped, output capped.
 */
export function renderTeamsMessageBody(message: GraphChatMessage): string {
  let content = text(message.body?.content);
  if (!content) return "";
  // A `text` body is already plain; only its length and whitespace need attention.
  if (text(message.body?.contentType)?.toLowerCase() === "text") return normalizeBody(content);
  return normalizeBody(renderMarkup(content, message));
}

// ── Search ──────────────────────────────────────────────────────────

/**
 * Validate an agent-supplied KQL search string for the search API.
 *
 * Double quotes are ALLOWED, unlike the mailbox validator: this query travels inside a JSON body,
 * where `JSON.stringify` is what keeps a value a value, and KQL phrase search (`"quarterly report"`)
 * is worth having. Control characters are refused because they cannot appear in a legitimate query
 * and would be echoed into an approval prompt; the length cap keeps a pathological query from
 * becoming a Graph error loop.
 */
export function validateTeamsSearchQuery(query: string): string {
  let trimmed = query.trim();
  if (!trimmed) throw new Error("Search query must not be empty.");
  if (trimmed.length > MAX_SEARCH_QUERY_CHARS) {
    throw new Error(`Search query must be at most ${MAX_SEARCH_QUERY_CHARS} characters.`);
  }
  // oxlint-disable-next-line no-control-regex -- intentionally rejecting control chars (echo guard)
  if (/[\u0000-\u001f\u007f]/.test(trimmed)) {
    throw new Error("Search query must not contain control characters.");
  }
  return trimmed;
}

/**
 * Build the `/search/query` request body.
 *
 * The one place a caller-supplied string reaches a Graph request body. The skeleton is literal: the
 * entity type, the page size, and the shape are written here, and `queryString` is the only value
 * that comes from outside — so a crafted query can populate `query.queryString` and nothing else,
 * regardless of what it contains. `from` is a number this file clamps, never a caller's string.
 */
export function buildSearchRequestBody(queryString: string, from: number): string {
  let offset = Number.isFinite(from) ? Math.min(Math.max(Math.trunc(from), 0), MAX_SEARCH_OFFSET) : 0;
  return JSON.stringify({
    requests: [{
      entityTypes: ["chatMessage"],
      query: { queryString },
      from: offset,
      size: SEARCH_PAGE_SIZE,
    }],
  });
}

/**
 * Whether a search failure is the index's result-window ceiling rather than a real error.
 *
 * Microsoft Search refuses to page past a fixed number of results and answers the request that
 * crosses the line with a 400. That is the end of the result set, not a failure, so the walk reports
 * it as exhaustion — otherwise a search with many matches would end in an error for every agent that
 * paged to the bottom.
 *
 * The caller checks that the walk has actually paged before consulting this: a first request starts
 * at offset zero and cannot cross a paging window, so a 400 there is a real error however its text
 * happens to read.
 */
function isSearchWindowExhausted(error: GraphApiError): boolean {
  if (error.status !== 400) return false;
  let detail = `${error.code} ${error.message}`;
  return /result window/i.test(detail)
      || /maximum (?:paging|offset|result|from)/i.test(detail)
      || /\bfrom\b[^.]{0,80}\bsize\b[^.]{0,80}(?:exceed|maximum|limit|less than)/i.test(detail);
}

/** The advance one `peekPage()` computed, applied by the `commitPage()` that follows it. */
type SearchWalkAdvance = {
  hits: TeamsSearchHitInfo[];
  offset: number;
  pages: number;
  exhausted: boolean;
  seen: Set<string>;
};

/**
 * A walk over search results.
 *
 * Search pages by offset, not by a link, so the continuation state lives here rather than in
 * anything a caller can hand back: a caller cannot ask for an arbitrary offset, only for the next
 * page. Two consequences of a live index are handled here too — the corpus shifts between requests,
 * so hits already returned are filtered out by id rather than shown twice; and the index stops
 * paging at a fixed window, which reads as "no more results" instead of an error.
 *
 * Reading a page is two calls, `peekPage()` then `commitPage()`, because a page has to be approved
 * before anyone may see it: the peek fetches and stages, the commit is what advances the walk. A
 * caller that peeks and then does not commit — a refused approval, an error on the way out — leaves
 * the walk exactly where it was, and its next peek answers with the same page rather than the one
 * after it. Callers that authorize between the two serialize the whole peek-authorize-commit
 * sequence themselves; the serialization here only keeps two peeks from spending two requests on
 * the same offset.
 */
export class TeamsSearchWalk {
  #api: GraphTeamsApi;
  #query: string;
  #offset = 0;
  #pages = 0;
  #exhausted = false;
  #seen = new Set<string>();
  #pending: SearchWalkAdvance | undefined;
  #tail: Promise<void> = Promise.resolve();

  constructor(api: GraphTeamsApi, query: string) {
    this.#api = api;
    this.#query = query;
  }

  /**
   * The next page of hits not already returned, or null once the search has nothing more. Does not
   * advance the walk — call `commitPage()` once the page has been handed to whoever asked for it.
   */
  peekPage(): Promise<TeamsSearchHitInfo[] | null> {
    // Serialized: two overlapping peeks would each fetch the same offset, and whichever page was
    // committed would leave the other one's offset never asked for by anyone.
    let result = this.#tail.then(() => this.#peekPage());
    this.#tail = result.then(() => undefined, () => undefined);
    return result;
  }

  /**
   * Record the page the last `peekPage()` returned as delivered, so the next peek moves past it.
   *
   * Nothing else advances the walk. Committing without a peek to commit does nothing: there is no
   * page whose delivery it could be recording.
   */
  commitPage(): void {
    let pending = this.#pending;
    if (!pending) return;
    this.#offset = pending.offset;
    this.#pages = pending.pages;
    this.#exhausted = pending.exhausted;
    this.#seen = pending.seen;
    this.#pending = undefined;
  }

  async #peekPage(): Promise<TeamsSearchHitInfo[] | null> {
    // A page that was fetched but never committed is answered from what it already holds, so a
    // caller still deciding about it neither spends a second request nor moves the walk.
    if (this.#pending) return this.#pending.hits;
    if (this.#exhausted) return null;

    // Continuation state is staged in locals and written back by `commitPage()`. Advancing on the
    // fetch instead would let a refused page be skipped: the next call would return the page AFTER
    // the one nobody was allowed to see.
    let offset = this.#offset;
    let pages = this.#pages;
    let seen = new Set(this.#seen);
    let exhausted = false;

    while (!exhausted) {
      if (pages >= MAX_SEARCH_PAGES) {
        throw new Error(
            `This search has already returned ${MAX_SEARCH_PAGES} pages. Narrow the query instead ` +
            "of paging further.");
      }
      let page: { hits: TeamsSearchHitInfo[]; moreAvailable: boolean };
      try {
        page = await this.#api.searchMessages(this.#query, offset);
      } catch (error) {
        if (offset > 0 && error instanceof GraphApiError && isSearchWindowExhausted(error)) {
          this.#exhausted = true;
          return null;
        }
        throw error;
      }
      pages++;
      // Offsets are absolute and independent of how many hits came back, so the next page starts
      // one page size further in whatever this one returned.
      offset += SEARCH_PAGE_SIZE;
      if (!page.moreAvailable) exhausted = true;

      let fresh = page.hits.filter(hit => !seen.has(hit.messageId));
      for (let hit of fresh) seen.add(hit.messageId);
      if (fresh.length > 0) {
        this.#pending = { hits: fresh, offset, pages, exhausted, seen };
        return fresh;
      }
    }

    // Nothing came back for anyone to be shown, so the end of the walk is recorded here: an empty
    // result is not a page a caller can be asked to approve, and there is nothing to commit.
    this.#offset = offset;
    this.#pages = pages;
    this.#seen = seen;
    this.#exhausted = true;
    return null;
  }
}

// ── Client ──────────────────────────────────────────────────────────

export type GraphTeamsApiOptions = {
  /** Called when Graph answers a claims challenge, so the account can be marked dead. */
  onCredentialsRejected?: CredentialsRejectedReporter;
};

/** `$top` as Graph wants it: a positive integer within the Teams collection ceiling. */
function top(size: number): string {
  return String(Math.min(Math.max(Math.trunc(size), 1), MAX_PAGE_SIZE));
}

export class GraphTeamsApi {
  #getAccessToken: AccessTokenProvider;
  #onCredentialsRejected: CredentialsRejectedReporter | undefined;

  constructor(getAccessToken: AccessTokenProvider, opts: GraphTeamsApiOptions = {}) {
    this.#getAccessToken = getAccessToken;
    this.#onCredentialsRejected = opts.onCredentialsRejected;
  }

  /** The throttling ceiling, as an error the caller can act on. */
  #retryAfterPolicy: RetryAfterPolicy = {
    maxWaitMs: MAX_RETRY_AFTER_MS,
    tooLong: (requestedMs: number) => new GraphApiError(429, "activityLimitReached",
        "Microsoft Graph is throttling this Teams connection and asked to be left alone for " +
        `${Math.round(requestedMs / 1000)} seconds. Try again after that.`),
  };

  /**
   * The one place a Teams Graph request is made.
   *
   * No `Prefer` header: the immutable-id and text-body preferences are Outlook options (see the file
   * header). What is set here is the auth policy and the throttling ceiling, so no call site can
   * make a Teams request without them.
   */
  async #request(url: string, init: RequestInit = {}): Promise<Response> {
    let headers = new Headers(init.headers);
    if (!headers.has("Accept")) headers.set("Accept", "application/json");
    if (init.body && !headers.has("Content-Type")) headers.set("Content-Type", "application/json");

    return await fetchWithAuthRetry(url, { ...init, headers }, this.#getAccessToken, {
      timeoutMs: GRAPH_TIMEOUT_MS,
      retryAfter: this.#retryAfterPolicy,
      ...(this.#onCredentialsRejected
          ? { onCredentialsRejected: this.#onCredentialsRejected }
          : {}),
    });
  }

  async #fetchJson<T>(url: string, init?: RequestInit): Promise<T> {
    let response = await this.#request(url, init);
    if (!response.ok) throw await this.#toError(response);
    if (response.status === 204) return undefined as T;
    return await response.json<T>();
  }

  /** Map a non-2xx Graph response onto an error stating what the caller can do about it. */
  async #toError(response: Response): Promise<GraphApiError> {
    let body = await response.json<{ error?: { code?: unknown; message?: unknown } }>()
        .catch(() => undefined);
    let code = typeof body?.error?.code === "string" ? body.error.code : "";
    let detail = truncate(
        typeof body?.error?.message === "string" && body.error.message
            ? body.error.message
            : `${response.status} ${response.statusText}`);

    if (response.status === 401) {
      let claims = claimsChallengeDetail(response.headers.get("WWW-Authenticate"));
      if (claims) {
        // fetchWithAuthRetry has already reported this to the account, which is what makes the
        // Workshop offer a reconnect. Nothing here retries.
        return new GraphApiError(401, code || claims,
            "Microsoft requires this account to sign in again before Teams can be used " +
            `(${claims}). Please reconnect the account.`,
            { credentialsRejected: true });
      }
      return new GraphApiError(401, code || "unauthenticated",
          `Microsoft rejected the Teams credentials (${detail}).`);
    }

    if (response.status === 403) {
      // Two of the Teams permissions require a tenant administrator's consent, so a 403 here is
      // as often "consent was never granted" as "this user cannot see it".
      return new GraphApiError(403, code || "forbidden",
          `This Microsoft connection is not permitted to read that Teams data (${detail}). ` +
          "Reconnect the account, and check that an administrator has granted the Teams " +
          "permissions.");
    }

    if (response.status === 404) {
      return new GraphApiError(404, code || "notFound",
          `That Teams item no longer exists, or this account cannot see it (${detail}).`);
    }

    if (response.status === 429) {
      let retryAfter = response.headers.get("Retry-After");
      return new GraphApiError(429, code || "activityLimitReached",
          "Microsoft Graph is throttling this Teams connection" +
          (retryAfter ? `; retry in ${truncate(retryAfter, 20)} seconds` : "") + ".");
    }

    if (response.status >= 500) {
      return new GraphApiError(response.status, code || "serviceError",
          `Microsoft Graph is temporarily unavailable (${detail}).`);
    }

    return new GraphApiError(response.status, code || "unknown",
        `Microsoft Graph request failed: ${response.status}${code ? ` ${code}` : ""} — ${detail}`);
  }

  /** One page of a collection, with its next link validated on arrival rather than when followed. */
  async #page<W, T>(url: string, map: (item: W) => T): Promise<GraphPage<T>> {
    let body = await this.#fetchJson<GraphCollection<W>>(url);
    let next = body["@odata.nextLink"];
    return {
      items: (body.value ?? []).map(map),
      ...(next ? { nextLink: assertGraphUrl(next) } : {}),
    };
  }

  /**
   * Every item of a collection the API returns whole, bounded by items and by requests.
   *
   * Reaching either bound with pages still outstanding is an error, not a short answer: a truncated
   * list is indistinguishable from a complete one, so a caller would go on to act on part of a team
   * believing it had all of it. `what` names the collection in that error.
   */
  async #collect<W, T>(url: string, map: (item: W) => T, maxItems: number, what: string)
      : Promise<T[]> {
    let items: T[] = [];
    let next: string = url;
    for (let page = 0; page < MAX_WALK_PAGES; page++) {
      let body: GraphCollection<W> = await this.#fetchJson<GraphCollection<W>>(next);
      for (let item of body.value ?? []) items.push(map(item));
      let link = body["@odata.nextLink"];
      if (!link) return items;
      next = assertGraphUrl(link);
      if (items.length >= maxItems) break;
    }
    throw new Error(
        `There are more ${what} than this can list in one go (the ceiling is ${maxItems}). Narrow ` +
        "what you are looking at — ask for a specific one by id rather than enumerating them all.");
  }

  // ── Teams and channels ────────────────────────────────────────────────────────

  /** The teams the connected user has joined. Throws rather than answering a truncated list. */
  async listJoinedTeams(): Promise<TeamsTeamInfo[]> {
    return await this.#collect<GraphTeam, TeamsTeamInfo>(
        graphUrl(["me", "joinedTeams"]), teamInfoFrom, MAX_TEAMS, "joined teams");
  }

  async getTeam(teamId: string): Promise<TeamsTeamInfo> {
    return teamInfoFrom(await this.#fetchJson<GraphTeam>(graphUrl(["teams", teamId])));
  }

  /**
   * The channels of a team that the connected user can see. Throws rather than answering a
   * truncated list.
   */
  async listChannels(teamId: string): Promise<TeamsChannelInfo[]> {
    return await this.#collect<GraphChannel, TeamsChannelInfo>(
        graphUrl(["teams", teamId, "channels"]),
        channel => channelInfoFrom(teamId, channel),
        MAX_CHANNELS,
        "channels in this team");
  }

  async getChannel(teamId: string, channelId: string): Promise<TeamsChannelInfo> {
    let channel = await this.#fetchJson<GraphChannel>(
        graphUrl(["teams", teamId, "channels", channelId]));
    return channelInfoFrom(teamId, channel);
  }

  async listTeamMembers(teamId: string): Promise<GraphPage<TeamsMemberInfo>> {
    return await this.#page<GraphConversationMember, TeamsMemberInfo>(
        graphUrl(["teams", teamId, "members"], { $top: top(MEMBERS_PAGE_SIZE) }), memberInfoFrom);
  }

  /**
   * First page of a channel's top-level messages.
   *
   * No `$orderby`: Graph sorts a channel by the last-modified time of each whole reply chain and
   * rejects any attempt to order it otherwise, so the order is Graph's and the agent-facing types
   * say so rather than pretending it is chronological.
   */
  async listChannelMessages(teamId: string, channelId: string)
      : Promise<GraphPage<TeamsMessageInfo>> {
    return await this.#messagePage(
        graphUrl(["teams", teamId, "channels", channelId, "messages"],
            { $top: top(TEAMS_MESSAGE_PAGE_SIZE) }));
  }

  /** First page of the replies under a channel message. */
  async listReplies(teamId: string, channelId: string, messageId: string)
      : Promise<GraphPage<TeamsMessageInfo>> {
    return await this.#messagePage(
        graphUrl(["teams", teamId, "channels", channelId, "messages", messageId, "replies"],
            { $top: top(TEAMS_MESSAGE_PAGE_SIZE) }));
  }

  async getChannelMessage(teamId: string, channelId: string, messageId: string)
      : Promise<TeamsMessageInfo> {
    return messageInfoFrom(await this.#fetchJson<GraphChatMessage>(
        graphUrl(["teams", teamId, "channels", channelId, "messages", messageId])));
  }

  // ── Chats ─────────────────────────────────────────────────────────────────────

  /**
   * First page of the connected user's chats.
   *
   * Deliberately without `$expand=members`: the expansion caps the members it inlines at 25 per chat
   * regardless of the page size asked for, which would both truncate rosters and invalidate this
   * page size. Rosters come from `listChatMembers`, which returns all of them.
   */
  async listChats(): Promise<GraphPage<TeamsChatInfo>> {
    return await this.#page<GraphChat, TeamsChatInfo>(
        graphUrl(["me", "chats"], { $top: top(CHATS_PAGE_SIZE) }), chatInfoFrom);
  }

  async getChat(chatId: string): Promise<TeamsChatInfo> {
    return chatInfoFrom(await this.#fetchJson<GraphChat>(graphUrl(["chats", chatId])));
  }

  /**
   * First page of a chat's members.
   *
   * No query parameters at all: this endpoint rejects them, so the page size is Graph's to choose
   * and the caller pages until the links run out.
   */
  async listChatMembers(chatId: string): Promise<GraphPage<TeamsMemberInfo>> {
    return await this.#page<GraphConversationMember, TeamsMemberInfo>(
        graphUrl(["chats", chatId, "members"]), memberInfoFrom);
  }

  /**
   * First page of a chat's messages, newest first by when they were sent.
   *
   * The order is requested explicitly. Graph's default for this collection is
   * `lastModifiedDateTime`, under which editing an old message pulls it back to the front and a page
   * walk can revisit ground it already covered; ordering by creation makes the walk monotonic.
   * Descending is the only direction Graph accepts here, and it will only filter on whatever
   * property it orders by — so any future `$filter` on this call has to be on `createdDateTime` too.
   */
  async listChatMessages(chatId: string): Promise<GraphPage<TeamsMessageInfo>> {
    return await this.#messagePage(graphUrl(["chats", chatId, "messages"], {
      $top: top(TEAMS_MESSAGE_PAGE_SIZE),
      $orderby: "createdDateTime desc",
    }));
  }

  async getChatMessage(chatId: string, messageId: string): Promise<TeamsMessageInfo> {
    return messageInfoFrom(await this.#fetchJson<GraphChatMessage>(
        graphUrl(["chats", chatId, "messages", messageId])));
  }

  // ── Pagination ────────────────────────────────────────────────────────────────
  //
  // A `@odata.nextLink` is followed verbatim — Graph has already put the page size, the ordering,
  // and its own continuation token in it, and re-adding any of them would either be rejected or
  // silently restart the walk.

  async nextMessagePage(nextLink: string): Promise<GraphPage<TeamsMessageInfo>> {
    return await this.#messagePage(assertGraphUrl(nextLink));
  }

  async nextMemberPage(nextLink: string): Promise<GraphPage<TeamsMemberInfo>> {
    return await this.#page<GraphConversationMember, TeamsMemberInfo>(
        assertGraphUrl(nextLink), memberInfoFrom);
  }

  async nextChatPage(nextLink: string): Promise<GraphPage<TeamsChatInfo>> {
    return await this.#page<GraphChat, TeamsChatInfo>(assertGraphUrl(nextLink), chatInfoFrom);
  }

  #messagePage(url: string): Promise<GraphPage<TeamsMessageInfo>> {
    return this.#page<GraphChatMessage, TeamsMessageInfo>(url, messageInfoFrom);
  }

  // ── Search ────────────────────────────────────────────────────────────────────

  /**
   * One page of message search results, starting at `from`.
   *
   * Hits are metadata only — the index stores no body, mentions, or attachments, and no team or
   * channel names — so they map to search-hit info and never to a message. `getSearchHitMessage`
   * turns one into a real message when the caller asks for it.
   */
  async searchMessages(query: string, from = 0)
      : Promise<{ hits: TeamsSearchHitInfo[]; moreAvailable: boolean }> {
    let validated = validateTeamsSearchQuery(query);
    let response = await this.#fetchJson<GraphSearchResponse>(graphUrl(["search", "query"]), {
      method: "POST",
      body: buildSearchRequestBody(validated, from),
    });
    let container = response.value?.[0]?.hitsContainers?.[0];
    let hits = (container?.hits ?? [])
        .map(searchHitInfoFrom)
        .filter((hit): hit is TeamsSearchHitInfo => hit !== undefined);
    return { hits, moreAvailable: container?.moreResultsAvailable === true };
  }

  /** A walk over every page of a search, validated up front so a bad query fails before any call. */
  openMessageSearch(query: string): TeamsSearchWalk {
    return new TeamsSearchWalk(this, validateTeamsSearchQuery(query));
  }

  /**
   * Read the message behind a search hit.
   *
   * A hit whose location the index did not report cannot be fetched from anywhere, so this fails
   * rather than returning an empty message: the hit was shown to the agent, and "this one cannot be
   * read" is the honest answer to asking for it.
   */
  async getSearchHitMessage(hit: TeamsSearchHitInfo): Promise<TeamsMessageInfo> {
    let location = hit.location;
    switch (location.kind) {
      case "channel":
        return await this.getChannelMessage(location.teamId, location.channelId, hit.messageId);
      case "chat":
        return await this.getChatMessage(location.chatId, hit.messageId);
      case "unknown":
        throw new GraphApiError(404, "unresolvedSearchHitLocation",
            "This search hit's location could not be resolved: Microsoft's search index reported " +
            "neither a channel nor a chat for it, so the message cannot be read.");
      default:
        location satisfies never;
        throw new Error(
            `unknown search hit location: ${(location as { kind: string }).kind}`);
    }
  }
}
