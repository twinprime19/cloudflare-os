// Microsoft Teams gatekeeper.
//
// Capability-based API over the Teams Graph client next door: `TeamsSession` hands out team,
// channel, chat, and message capabilities, and cursors pull pages as the gadget consumes them. Each
// session read is authorized through the approval queue before anything is returned to the caller;
// the agent catalog is the one surface outside that gate (see getAgentCatalog).
//
// Read-only, structurally. There is no action surface at all — no pending-action store, no
// `submitAction` call site, and `applyAction`/`rejectAction`/`revertAction` throw rather than
// looking anything up, because no action id can ever exist. Nothing here posts, edits, deletes,
// marks read, joins, or leaves; the one non-GET request this resource makes is the search POST
// inside the client, which is a query.
//
// Echo discipline. Unlike the mailbox next door, whose folder names the account owner chose,
// nearly every string Teams returns was written by somebody else: team and channel names, chat
// topics, display names, message bodies. So the two echo slots are treated differently, everywhere:
//
//   - A title is one line in an approval list and cannot carry a fence, so it is a literal plus at
//     most one name passed through `sanitizeApprovalTitle`. A message body never appears in one.
//   - A description carries the counterpart text, and every such value goes through
//     `formatApprovalField`, which fences it so it cannot forge the Markdown around it.
//
// A cursor's scope follows the same rule: the phrase naming what is being paged ("the selected
// channel") is a literal written here, and the name it refers to is fenced separately in the same
// description rather than interpolated into the phrase.
//
// Errors name ids, never counterpart display text — a thrown message is rendered in places an
// approval description is not.

import { DurableObject, RpcStub, RpcTarget } from "cloudflare:workers";
import { validateRpc } from "capnweb-validate";
import {
  ActionKind, AgentCatalog, ApprovalQueue, Cursor, Gatekeeper, GatekeeperUserVerifier,
  ResourceDescription, boundAgentCatalog,
} from "@gadgets/workshop-shared/gatekeeper";
import { formatApprovalField, sanitizeApprovalTitle } from "./approval-text";
import { AccessTokenCache, AccessTokenRequest } from "./auth-retry";
import type { GraphPage } from "./graph-api";
import { GraphTeamsApi, TeamsSearchWalk, validateTeamsSearchQuery } from "./graph-teams-api";
import type {
  TeamsChannel, TeamsChannelEntry, TeamsChannelInfo, TeamsChannelMessage, TeamsChat,
  TeamsChatEntry, TeamsChatInfo, TeamsMemberInfo, TeamsMessage, TeamsMessageEntry,
  TeamsMessageInfo, TeamsSearchHitEntry, TeamsSearchHitInfo, TeamsSession, TeamsTeam,
  TeamsTeamEntry, TeamsTeamInfo, TeamsUser,
} from "./teams-types";
import TEAMS_TYPES_CODE from "./teams-types.txt";
import type { Env, UserAccount } from "./microsoft";

/**
 * Canonical URL of the Teams resource. A connected account has exactly one Teams surface, so this
 * is a constant rather than something the configurator chooses.
 */
export const TEAMS_URL = "https://teams.microsoft.com/";

/** Ceiling on pages a single cursor will pull, so an agent cannot walk a whole history. */
const MAX_CURSOR_PAGES = 40;

/** Consecutive empty-but-continuing pages a cursor skips before reporting exhaustion. */
const MAX_EMPTY_PAGES = 5;

/** Longest excerpt of a message body placed in an approval description. */
const MAX_EXCERPT_CHARS = 200;

/** Longest excerpt of one message on a page listing, where many share the prompt. */
const MAX_LINE_EXCERPT_CHARS = 120;

/** Longest display name echoed into a per-line entry of a page listing. */
const MAX_NAME_CHARS = 80;

/**
 * Combined ceiling on the teams and chats offered to the agent catalog. Both grow with the user's
 * membership, so they share a bound well under the shared AGENT_CATALOG_MAX_ENTRIES ceiling;
 * everything past it stays reachable through the session's listTeams() and listChats().
 */
const MAX_CATALOG_ENTRIES = 25;

// ── Session context ─────────────────────────────────────────────────
//
// No pending-action store, unlike the mailbox: this resource queues nothing, so there is no state
// an approval decision could apply to.

type TeamsSessionContext = {
  api: GraphTeamsApi;
  approvalQueue: RpcStub<ApprovalQueue>;
};

/** One bounded line standing for a value that may be long or many-lined. */
function oneLine(value: string, max: number): string {
  let flat = value.replace(/\s+/g, " ").trim();
  return flat.length > max ? `${flat.slice(0, max)}…` : flat;
}

function senderName(from: TeamsUser | undefined): string {
  return from?.displayName ?? "(unknown sender)";
}

/**
 * What to call a chat. Most one-on-one chats and plenty of group chats have no topic, so the kind
 * stands in — a neutral literal, unlike the topic beside it.
 */
function chatLabel(info: TeamsChatInfo): string {
  return info.topic ?? `(${info.chatKind} chat)`;
}

/**
 * What a cursor is paging over, for the observation on each of its pages.
 *
 * `label` is a phrase written in this file. `field`, when present, is the counterpart-controlled
 * name of the thing that phrase refers to, which the description fences separately rather than
 * splicing into the phrase.
 */
type TeamsScope = {
  label: string;
  field?: { label: string; value: string };
};

function scopeOf(label: string, fieldLabel: string, value: string | undefined): TeamsScope {
  return value === undefined ? { label } : { label, field: { label: fieldLabel, value } };
}

/** The scope's name as a fenced field, ready to append to a description, or nothing. */
function scopeField(scope: TeamsScope): string {
  return scope.field ? `\n\n${formatApprovalField(scope.field.label, scope.field.value)}` : "";
}

function describeMessage(info: TeamsMessageInfo): string {
  return [
    formatApprovalField("From", senderName(info.from)),
    formatApprovalField("Posted", info.createdAt.toISOString()),
    formatApprovalField("Message", oneLine(info.body, MAX_EXCERPT_CHARS) || "(no text)"),
  ].join("\n\n");
}

/** One line per message — who wrote it, and the start of what they wrote. */
function describeMessageLines(items: TeamsMessageInfo[]): string {
  return items.map(info =>
      `${oneLine(senderName(info.from), MAX_NAME_CHARS)}: ` +
      `${oneLine(info.body, MAX_LINE_EXCERPT_CHARS) || "(no text)"}`).join("\n");
}

/** One line per search hit. The index reports no body, so the subject or the extract stands in. */
function describeHitLines(hits: TeamsSearchHitInfo[]): string {
  return hits.map(hit =>
      `${oneLine(senderName(hit.from), MAX_NAME_CHARS)}: ` +
      `${oneLine(hit.subject ?? hit.summary, MAX_LINE_EXCERPT_CHARS) || "(no text)"}`).join("\n");
}

// ── Paged cursors ───────────────────────────────────────────────────
// Lazily pull pages from Graph as the gadget calls next(), following `@odata.nextLink`. A cursor is
// itself a capability: the call that creates it authorizes its creation, and each next() separately
// authorizes the page it returns.
//
// The walk below is shared by messages, members, and chats, which differ only in what an item
// becomes and how a page is described. The part that has to be right — staging the continuation
// until the observation is authorized — is exactly the part that must not be written three times.
// The three `RpcTarget` classes after it are thin, and exist because the RPC argument validator
// needs to see the concrete shape a `next()` returns; a type parameter gives it nothing to check.

/** What one page of a cursor is reported as, once its entries exist. */
type PageObservation = { title: string; description: string };

type PageCursorSpec<T, E> = {
  /** The first request. Bound by whoever built the cursor; takes no continuation. */
  firstPage: () => Promise<GraphPage<T>>;
  /**
   * Follow a `@odata.nextLink`. The link is only ever one this walk read off a page it already
   * fetched: no agent-reachable value reaches this parameter, so a Graph continuation is never
   * something an agent can shape.
   */
  nextPage: (nextLink: string) => Promise<GraphPage<T>>;
  toEntries: (items: T[]) => E[];
  describePage: (entries: E[]) => PageObservation;
};

/** The staged walk over a link-paged Graph collection. Not itself exposed over RPC. */
class TeamsPageWalk<T, E> {
  #ctx: TeamsSessionContext;
  #spec: PageCursorSpec<T, E>;
  #nextLink: string | undefined;
  #started = false;
  #exhausted = false;
  #pages = 0;
  #tail: Promise<void> = Promise.resolve();

  constructor(ctx: TeamsSessionContext, spec: PageCursorSpec<T, E>) {
    this.#ctx = ctx;
    this.#spec = spec;
  }

  next(): Promise<E[] | null> {
    // Serialized: two overlapping next() calls would both read `#nextLink` before either advanced
    // it, returning the same page twice.
    let result = this.#tail.then(() => this.#nextPage());
    this.#tail = result.then(() => undefined, () => undefined);
    return result;
  }

  async #nextPage(): Promise<E[] | null> {
    if (this.#exhausted) return null;
    if (this.#pages >= MAX_CURSOR_PAGES) {
      throw new Error(
          `This cursor has already returned ${MAX_CURSOR_PAGES} pages. Narrow what you are ` +
          "reading instead of paging further.");
    }

    // Pagination state is staged in locals and only written back once the observation has been
    // authorized. Committing first would let a refused (or failed) authorization advance the
    // cursor, so the retry the caller makes would silently return the page AFTER the one it never
    // received.
    let page: GraphPage<T>;
    let nextLink = this.#nextLink;
    let started = this.#started;
    let pages = this.#pages;
    let skipped = 0;
    do {
      if (!started) {
        page = await this.#spec.firstPage();
        started = true;
      } else if (nextLink) {
        page = await this.#spec.nextPage(nextLink);
      } else {
        this.#exhausted = true;
        return null;
      }
      nextLink = page.nextLink;
      pages++;
      // Graph can answer with an empty page that still carries a next link (e.g. everything on it
      // was filtered out server-side). Skip those rather than reporting exhaustion.
      skipped++;
    } while (page.items.length === 0 && page.nextLink && skipped < MAX_EMPTY_PAGES &&
             pages < MAX_CURSOR_PAGES);

    if (page.items.length === 0) {
      // Nothing was returned to the caller, so only the terminal state is recorded — an empty page
      // is not a page anyone can be asked to re-read.
      this.#started = started;
      this.#nextLink = nextLink;
      this.#pages = pages;
      this.#exhausted = !page.nextLink;
      if (page.nextLink) {
        // A bound stopped the skip above while Graph still had more to give. `null` here would
        // read as "the listing is finished", so the caller is told a ceiling was hit instead —
        // the same promise the cursor makes when it has served its page ceiling.
        throw new Error(pages >= MAX_CURSOR_PAGES
            ? `This cursor has already read ${MAX_CURSOR_PAGES} pages. Narrow what you are ` +
              "reading instead of paging further."
            : `This cursor skipped ${MAX_EMPTY_PAGES} pages with nothing on them without ` +
              "reaching the end of the listing. Narrow what you are reading instead of paging " +
              "further.");
      }
      return null;
    }

    let entries = this.#spec.toEntries(page.items);
    await this.#ctx.approvalQueue.authorizeObservation(this.#spec.describePage(entries));

    this.#started = started;
    this.#nextLink = nextLink;
    this.#pages = pages;
    if (!page.nextLink) this.#exhausted = true;
    return entries;
  }
}

@validateRpc()
class TeamsMessageCursorImpl extends RpcTarget implements Cursor<TeamsMessageEntry> {
  #walk: TeamsPageWalk<TeamsMessageInfo, TeamsMessageEntry>;

  constructor(walk: TeamsPageWalk<TeamsMessageInfo, TeamsMessageEntry>) {
    super();
    this.#walk = walk;
  }

  next(): Promise<TeamsMessageEntry[] | null> {
    return this.#walk.next();
  }
}

@validateRpc()
class TeamsMemberCursorImpl extends RpcTarget implements Cursor<TeamsMemberInfo> {
  #walk: TeamsPageWalk<TeamsMemberInfo, TeamsMemberInfo>;

  constructor(walk: TeamsPageWalk<TeamsMemberInfo, TeamsMemberInfo>) {
    super();
    this.#walk = walk;
  }

  next(): Promise<TeamsMemberInfo[] | null> {
    return this.#walk.next();
  }
}

@validateRpc()
class TeamsChatCursorImpl extends RpcTarget implements Cursor<TeamsChatEntry> {
  #walk: TeamsPageWalk<TeamsChatInfo, TeamsChatEntry>;

  constructor(walk: TeamsPageWalk<TeamsChatInfo, TeamsChatEntry>) {
    super();
    this.#walk = walk;
  }

  next(): Promise<TeamsChatEntry[] | null> {
    return this.#walk.next();
  }
}

function messageCursor(
    ctx: TeamsSessionContext,
    scope: TeamsScope,
    firstPage: () => Promise<GraphPage<TeamsMessageInfo>>,
    toMessage?: (info: TeamsMessageInfo) => TeamsChannelMessage): Cursor<TeamsMessageEntry> {
  return new TeamsMessageCursorImpl(new TeamsPageWalk<TeamsMessageInfo, TeamsMessageEntry>(ctx, {
    firstPage,
    nextPage: nextLink => ctx.api.nextMessagePage(nextLink),
    toEntries: items => items.map(info => ({
      info,
      ...(toMessage ? { message: toMessage(info) } : {}),
    })),
    describePage: entries => ({
      title: `Read ${entries.length} Microsoft Teams messages`,
      description:
          `Fetch the next page of messages from ${scope.label}.${scopeField(scope)}\n\n` +
          formatApprovalField("Messages", describeMessageLines(entries.map(entry => entry.info))),
    }),
  }));
}

function memberCursor(
    ctx: TeamsSessionContext,
    scope: TeamsScope,
    firstPage: () => Promise<GraphPage<TeamsMemberInfo>>): Cursor<TeamsMemberInfo> {
  return new TeamsMemberCursorImpl(new TeamsPageWalk<TeamsMemberInfo, TeamsMemberInfo>(ctx, {
    firstPage,
    nextPage: nextLink => ctx.api.nextMemberPage(nextLink),
    toEntries: items => items,
    describePage: entries => ({
      title: `Read ${entries.length} Microsoft Teams members`,
      description:
          `Fetch the next page of the roster of ${scope.label}.${scopeField(scope)}\n\n` +
          formatApprovalField("Members",
              entries.map(entry => oneLine(entry.displayName, MAX_NAME_CHARS)).join("\n")),
    }),
  }));
}

function chatCursor(ctx: TeamsSessionContext): Cursor<TeamsChatEntry> {
  return new TeamsChatCursorImpl(new TeamsPageWalk<TeamsChatInfo, TeamsChatEntry>(ctx, {
    firstPage: () => ctx.api.listChats(),
    nextPage: nextLink => ctx.api.nextChatPage(nextLink),
    toEntries: items => items.map(info => ({
      info,
      chat: new TeamsChatStub(ctx, info.id, info) as TeamsChat,
    })),
    describePage: entries => ({
      title: `Read ${entries.length} Microsoft Teams chats`,
      description:
          "Fetch the next page of the connected user's Teams chats.\n\n" +
          formatApprovalField("Chats",
              entries.map(entry => oneLine(chatLabel(entry.info), MAX_NAME_CHARS)).join("\n")),
    }),
  }));
}

// ── TeamsSearchCursorImpl ───────────────────────────────────────────

@validateRpc()
class TeamsSearchCursorImpl extends RpcTarget implements Cursor<TeamsSearchHitEntry> {
  #ctx: TeamsSessionContext;
  #query: string;
  #walk: TeamsSearchWalk;
  #tail: Promise<void> = Promise.resolve();

  constructor(ctx: TeamsSessionContext, query: string, walk: TeamsSearchWalk) {
    super();
    this.#ctx = ctx;
    this.#query = query;
    this.#walk = walk;
  }

  next(): Promise<TeamsSearchHitEntry[] | null> {
    let result = this.#tail.then(() => this.#nextPage());
    this.#tail = result.then(() => undefined, () => undefined);
    return result;
  }

  /**
   * Peek, authorize, commit — the whole sequence inside the serialized section.
   *
   * The walk stages a peeked page and moves past it only on `commitPage()`, and a commit with no
   * staged page does nothing. So the three steps have to belong to one caller: were a second
   * next() to peek while this one was still waiting for its authorization, it would be handed the
   * same staged page, and whichever commit ran second would silently do nothing — the same page
   * delivered twice, and the walk advanced once.
   */
  async #nextPage(): Promise<TeamsSearchHitEntry[] | null> {
    let hits = await this.#walk.peekPage();
    if (!hits) return null;

    let entries = hits.map(info => ({
      info,
      message: new TeamsSearchHitStub(this.#ctx, info) as TeamsMessage,
    }));

    await this.#ctx.approvalQueue.authorizeObservation({
      title: `Read ${entries.length} Microsoft Teams search results`,
      description:
          "Fetch the next page of Teams messages matching this search.\n\n" +
          `${formatApprovalField("Query", this.#query)}\n\n` +
          formatApprovalField("Results", describeHitLines(hits)),
    });

    // Only now does the walk move past this page. A refused authorization throws above, leaving the
    // page staged, so the caller's retry is answered with the same page rather than the one after.
    this.#walk.commitPage();
    return entries;
  }
}

// ── TeamsTeamStub ───────────────────────────────────────────────────

@validateRpc()
class TeamsTeamStub extends RpcTarget implements TeamsTeam {
  #ctx: TeamsSessionContext;
  #teamId: string;
  #cachedInfo: TeamsTeamInfo | undefined;

  constructor(ctx: TeamsSessionContext, teamId: string, cachedInfo?: TeamsTeamInfo) {
    super();
    this.#ctx = ctx;
    this.#teamId = teamId;
    this.#cachedInfo = cachedInfo;
  }

  async getInfo(): Promise<TeamsTeamInfo> {
    let info = this.#cachedInfo ?? await this.#ctx.api.getTeam(this.#teamId);
    this.#cachedInfo = info;

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Teams team: ${info.displayName}`),
      description:
          "Read metadata for this Microsoft Teams team.\n\n" +
          formatApprovalField("Team", info.displayName),
    });

    return info;
  }

  async listChannels(): Promise<TeamsChannelEntry[]> {
    let channels = await this.#ctx.api.listChannels(this.#teamId);
    let scope = this.#scope();

    await this.#ctx.approvalQueue.authorizeObservation({
      title: `List ${channels.length} Microsoft Teams channels`,
      description:
          `List the channels of ${scope.label} that the connected user can see.` +
          `${scopeField(scope)}\n\n` +
          formatApprovalField("Channels",
              channels.map(info => oneLine(info.displayName, MAX_NAME_CHARS)).join("\n")),
    });

    return channels.map(info => ({
      info,
      channel: new TeamsChannelStub(this.#ctx, this.#teamId, info.id, info) as TeamsChannel,
    }));
  }

  /**
   * A capability for one channel, read before it is handed over: the caller is told this throws
   * when the channel cannot be seen, and the channel's name is what makes the observation
   * reviewable at all.
   */
  async getChannel(channelId: string): Promise<TeamsChannel> {
    let info = await this.#ctx.api.getChannel(this.#teamId, channelId);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Teams channel: ${info.displayName}`),
      description:
          "Open a Microsoft Teams channel by id.\n\n" +
          formatApprovalField("Channel", info.displayName),
    });

    return new TeamsChannelStub(this.#ctx, this.#teamId, info.id, info);
  }

  async listMembers(): Promise<Cursor<TeamsMemberInfo>> {
    let scope = this.#scope();

    await this.#ctx.approvalQueue.authorizeObservation({
      title: "List Microsoft Teams team members",
      description: `Create a cursor over the roster of ${scope.label}.${scopeField(scope)}`,
    });

    return memberCursor(this.#ctx, scope, () => this.#ctx.api.listTeamMembers(this.#teamId));
  }

  #scope(): TeamsScope {
    return scopeOf("the selected team", "Team", this.#cachedInfo?.displayName);
  }
}

// ── TeamsChannelStub ────────────────────────────────────────────────

@validateRpc()
class TeamsChannelStub extends RpcTarget implements TeamsChannel {
  #ctx: TeamsSessionContext;
  #teamId: string;
  #channelId: string;
  #cachedInfo: TeamsChannelInfo | undefined;

  constructor(
      ctx: TeamsSessionContext, teamId: string, channelId: string,
      cachedInfo?: TeamsChannelInfo) {
    super();
    this.#ctx = ctx;
    this.#teamId = teamId;
    this.#channelId = channelId;
    this.#cachedInfo = cachedInfo;
  }

  async getInfo(): Promise<TeamsChannelInfo> {
    let info = this.#cachedInfo ?? await this.#ctx.api.getChannel(this.#teamId, this.#channelId);
    this.#cachedInfo = info;

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Teams channel: ${info.displayName}`),
      description:
          "Read metadata for this Microsoft Teams channel.\n\n" +
          formatApprovalField("Channel", info.displayName),
    });

    return info;
  }

  async listMessages(): Promise<Cursor<TeamsMessageEntry>> {
    let scope = scopeOf("the selected channel", "Channel", this.#cachedInfo?.displayName);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: "List Microsoft Teams channel messages",
      description:
          `Create a cursor over the top-level messages of ${scope.label}.${scopeField(scope)}`,
    });

    return messageCursor(
        this.#ctx, scope,
        () => this.#ctx.api.listChannelMessages(this.#teamId, this.#channelId),
        // A channel's top-level messages are the roots of its reply chains, so each one carries a
        // capability that can read them. Replies and chat messages carry none: neither can be
        // replied to in turn.
        info => new TeamsChannelMessageStub(
            this.#ctx, this.#teamId, this.#channelId, info.id, info) as TeamsChannelMessage);
  }
}

// ── TeamsChatStub ───────────────────────────────────────────────────

@validateRpc()
class TeamsChatStub extends RpcTarget implements TeamsChat {
  #ctx: TeamsSessionContext;
  #chatId: string;
  #cachedInfo: TeamsChatInfo | undefined;

  constructor(ctx: TeamsSessionContext, chatId: string, cachedInfo?: TeamsChatInfo) {
    super();
    this.#ctx = ctx;
    this.#chatId = chatId;
    this.#cachedInfo = cachedInfo;
  }

  async getInfo(): Promise<TeamsChatInfo> {
    let info = this.#cachedInfo ?? await this.#ctx.api.getChat(this.#chatId);
    this.#cachedInfo = info;

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Teams chat: ${chatLabel(info)}`),
      description:
          "Read metadata for this Microsoft Teams chat.\n\n" +
          formatApprovalField("Chat", chatLabel(info)),
    });

    return info;
  }

  async listMembers(): Promise<Cursor<TeamsMemberInfo>> {
    let scope = this.#scope();

    await this.#ctx.approvalQueue.authorizeObservation({
      title: "List Microsoft Teams chat members",
      description: `Create a cursor over everyone in ${scope.label}.${scopeField(scope)}`,
    });

    return memberCursor(this.#ctx, scope, () => this.#ctx.api.listChatMembers(this.#chatId));
  }

  async listMessages(): Promise<Cursor<TeamsMessageEntry>> {
    let scope = this.#scope();

    await this.#ctx.approvalQueue.authorizeObservation({
      title: "List Microsoft Teams chat messages",
      description: `Create a cursor over the messages in ${scope.label}.${scopeField(scope)}`,
    });

    return messageCursor(this.#ctx, scope, () => this.#ctx.api.listChatMessages(this.#chatId));
  }

  #scope(): TeamsScope {
    return scopeOf(
        "the selected chat", "Chat", this.#cachedInfo ? chatLabel(this.#cachedInfo) : undefined);
  }
}

// ── TeamsChannelMessageStub ─────────────────────────────────────────

@validateRpc()
class TeamsChannelMessageStub extends RpcTarget implements TeamsChannelMessage {
  #ctx: TeamsSessionContext;
  #teamId: string;
  #channelId: string;
  #messageId: string;
  #cachedInfo: TeamsMessageInfo | undefined;

  constructor(
      ctx: TeamsSessionContext, teamId: string, channelId: string, messageId: string,
      cachedInfo?: TeamsMessageInfo) {
    super();
    this.#ctx = ctx;
    this.#teamId = teamId;
    this.#channelId = channelId;
    this.#messageId = messageId;
    this.#cachedInfo = cachedInfo;
  }

  /**
   * Always re-read, even when this stub came from a listing that already carried the message. A
   * Teams message can be edited or deleted after it was listed, and a stale answer would be
   * indistinguishable from a fresh one; callers that want the listed copy already have it on the
   * entry's `info`.
   */
  async getInfo(): Promise<TeamsMessageInfo> {
    let info = await this.#ctx.api.getChannelMessage(
        this.#teamId, this.#channelId, this.#messageId);
    this.#cachedInfo = info;

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Teams message from ${senderName(info.from)}`),
      description: `Read a Microsoft Teams channel message.\n\n${describeMessage(info)}`,
    });

    return info;
  }

  async listReplies(): Promise<Cursor<TeamsMessageEntry>> {
    let body = this.#cachedInfo?.body;
    let scope = scopeOf(
        "the reply chain under the selected message", "Message",
        body === undefined ? undefined : oneLine(body, MAX_EXCERPT_CHARS) || "(no text)");

    await this.#ctx.approvalQueue.authorizeObservation({
      title: "List Microsoft Teams message replies",
      description: `Create a cursor over ${scope.label}.${scopeField(scope)}`,
    });

    return messageCursor(this.#ctx, scope,
        () => this.#ctx.api.listReplies(this.#teamId, this.#channelId, this.#messageId));
  }
}

// ── TeamsSearchHitStub ──────────────────────────────────────────────
//
// The message behind a search hit. Deliberately only a `TeamsMessage`: a hit may name a chat
// message as readily as a channel one, and a chat message has no reply chain to offer.

@validateRpc()
class TeamsSearchHitStub extends RpcTarget implements TeamsMessage {
  #ctx: TeamsSessionContext;
  #hit: TeamsSearchHitInfo;

  constructor(ctx: TeamsSessionContext, hit: TeamsSearchHitInfo) {
    super();
    this.#ctx = ctx;
    this.#hit = hit;
  }

  async getInfo(): Promise<TeamsMessageInfo> {
    let info = await this.#ctx.api.getSearchHitMessage(this.#hit);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Teams message from ${senderName(info.from)}`),
      description:
          `Read the Microsoft Teams message behind a search result.\n\n${describeMessage(info)}`,
    });

    return info;
  }
}

// ── TeamsSessionImpl ────────────────────────────────────────────────

@validateRpc()
class TeamsSessionImpl extends RpcTarget implements TeamsSession {
  #ctx: TeamsSessionContext;

  constructor(ctx: TeamsSessionContext) {
    super();
    this.#ctx = ctx;
  }

  async listTeams(): Promise<TeamsTeamEntry[]> {
    let teams = await this.#ctx.api.listJoinedTeams();

    await this.#ctx.approvalQueue.authorizeObservation({
      title: `List ${teams.length} Microsoft Teams teams`,
      description:
          "List the teams the connected user has joined.\n\n" +
          formatApprovalField("Teams",
              teams.map(info => oneLine(info.displayName, MAX_NAME_CHARS)).join("\n")),
    });

    return teams.map(info => ({
      info,
      team: new TeamsTeamStub(this.#ctx, info.id, info) as TeamsTeam,
    }));
  }

  async getTeam(teamId: string): Promise<TeamsTeam> {
    let info = await this.#ctx.api.getTeam(teamId);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Teams team: ${info.displayName}`),
      description:
          "Open a Microsoft Teams team by id.\n\n" +
          formatApprovalField("Team", info.displayName),
    });

    return new TeamsTeamStub(this.#ctx, info.id, info);
  }

  async listChats(): Promise<Cursor<TeamsChatEntry>> {
    await this.#ctx.approvalQueue.authorizeObservation({
      title: "List Microsoft Teams chats",
      description: "Create a cursor over the chats the connected user takes part in.",
    });

    return chatCursor(this.#ctx);
  }

  async getChat(chatId: string): Promise<TeamsChat> {
    let info = await this.#ctx.api.getChat(chatId);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Teams chat: ${chatLabel(info)}`),
      description:
          "Open a Microsoft Teams chat by id.\n\n" +
          formatApprovalField("Chat", chatLabel(info)),
    });

    return new TeamsChatStub(this.#ctx, info.id, info);
  }

  async search(query: string): Promise<Cursor<TeamsSearchHitEntry>> {
    // Validated before the approval prompt, so a malformed query fails immediately instead of
    // asking a human to approve a search that cannot run.
    let validated = validateTeamsSearchQuery(query);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: "Search Microsoft Teams",
      description:
          "Create a cursor over the Teams messages matching this search.\n\n" +
          formatApprovalField("Query", validated),
    });

    return new TeamsSearchCursorImpl(
        this.#ctx, validated, this.#ctx.api.openMessageSearch(validated));
  }
}

// =======================================================================================

export type TeamsGatekeeperImplProps = {
  userObjectId: string;
}

@validateRpc()
export class TeamsGatekeeperImpl
    extends DurableObject<Env, TeamsGatekeeperImplProps>
    implements Gatekeeper<TeamsSession> {
  #tokens = new AccessTokenCache(opts => this.#account().getAccessToken(opts));

  #account(): DurableObjectStub<UserAccount> {
    return this.ctx.exports.UserAccount.get(
        this.ctx.exports.UserAccount.idFromString(this.ctx.props.userObjectId));
  }

  #getAccessToken(opts?: AccessTokenRequest): Promise<string> {
    return this.#tokens.get(opts);
  }

  /**
   * A Graph client wired to this account.
   *
   * The claims-challenge hook reports straight to the account authority, which drops the cached
   * token and tells the Workshop the credentials are dead — the same path a permanently failed
   * token mint takes, so a Conditional Access rejection surfaces a reconnect prompt instead of an
   * account that refreshes happily while every Teams call fails.
   *
   * This object's own token memo is dropped first. It is the only copy the account authority cannot
   * reach, and it outlives the reconnect that fixes the account: keeping it would serve the rejected
   * token again after the user has reconnected, re-reporting a death that no longer exists.
   */
  #api(): GraphTeamsApi {
    return new GraphTeamsApi(opts => this.#getAccessToken(opts), {
      onCredentialsRejected: async (detail: string) => {
        this.#tokens.invalidate();
        await this.#account().reportCredentialsRejected(detail);
      },
    });
  }

  async describe(): Promise<ResourceDescription> {
    return {
      url: TEAMS_URL,
      title: "Microsoft Teams",
      snippet: "Your Microsoft Teams messages, read-only",
      suggestedBindingName: "MICROSOFT_TEAMS",
      tsType: "TeamsSession",
    };
  }

  async getTypeScriptTypes(): Promise<string> {
    return TEAMS_TYPES_CODE;
  }

  async getAutoApprovableActions(): Promise<ActionKind[]> {
    // This resource submits no actions at all, so there is no kind of action to auto-approve.
    return [];
  }

  async startSession(approvalQueue: RpcStub<ApprovalQueue>): Promise<TeamsSession> {
    // `dup()` because Cap'n Web disposes every stub passed as an argument once the call returns,
    // and the session keeps calling this queue long after startSession() has answered.
    return new TeamsSessionImpl({
      api: this.#api(),
      approvalQueue: approvalQueue.dup(),
    });
  }

  /**
   * Teams and chats by name, so an agent can find where a conversation lives without paging the
   * session API. Teams come first: they are the named, stable half of the surface, while a chat's
   * identity is often only who is in it.
   *
   * The Workshop loads this into every chat's prompt on every turn and no approval stands in the
   * way, so the titles — counterpart-controlled text, like nearly everything Teams returns — are
   * each flattened to a single line by `catalogTitle` and cannot forge the structure they land in.
   */
  async getAgentCatalog(): Promise<AgentCatalog | null> {
    let api = this.#api();
    let [teams, chats] = await Promise.all([api.listJoinedTeams(), api.listChats()]);

    let entries = [
      ...teams.map(team => ({
        id: team.id,
        title: catalogTitle(team.displayName),
        description: "Microsoft Teams team the connected user has joined. Open it with getTeam().",
      })),
      ...chats.items.map(chat => ({
        id: chat.id,
        title: catalogTitle(chatLabel(chat)),
        description:
            `Microsoft Teams ${chat.chatKind} chat. Open it with getChat().` +
            (chat.lastUpdatedAt ? ` Last active ${chat.lastUpdatedAt.toISOString()}.` : ""),
      })),
    ];

    let catalog = boundAgentCatalog(entries.slice(0, MAX_CATALOG_ENTRIES));
    // A user can belong to more teams and chats than the catalog advertises; boundAgentCatalog only
    // flags its own, far larger ceiling, so the drop at MAX_CATALOG_ENTRIES — and the chats past
    // the first page — are reported here.
    if (entries.length > MAX_CATALOG_ENTRIES || chats.nextLink) catalog.truncated = true;

    return catalog;
  }

  // ---------------------------------------------------------------------------
  // No action surface. Nothing in this gatekeeper submits an action, so no action id can exist and
  // there is no storage for one of these to look in. They are here because the contract requires
  // them, and they throw rather than pretending to have done something.

  async applyAction(actionId: number): Promise<void> {
    throw new Error(`Unknown pending Teams action: ${actionId}`);
  }

  async rejectAction(actionId: number): Promise<void | {restart?: boolean}> {
    throw new Error(`Unknown pending Teams action: ${actionId}`);
  }

  async revertAction(_action: number):
      Promise<void | {message?: string, canRetry?: boolean, restart?: boolean}> {
    throw new Error("revert is not implemented");
  }

  /**
   * Observer tracking — owner-only in this version. A binding here covers every team, channel, and
   * private chat the connected user takes part in, and no other user can be shown to have access to
   * all of that, so addObserver always throws and a Teams-bound gadget stays single-user.
   * removeObserver is a no-op since no observer is ever recorded.
   */
  async addObserver(_id: string, _user: Fetcher<GatekeeperUserVerifier>): Promise<void> {
    throw new Error(
        "Microsoft Teams data cannot be shared with other users: this connection reads the teams, " +
        "channels, and private chats of the account that connected it, so in this version it may " +
        "only be observed by that account's owner.");
  }

  async removeObserver(_id: string): Promise<void> {}
}

/** Flatten a name for the catalog. `boundAgentCatalog` applies the length cap. */
function catalogTitle(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}
