// Outlook mailbox gatekeeper.
//
// Capability-based API: OutlookMailSession returns folder and message cursors, whose entries carry
// stubs the gadget can hold and call later. Reads go through `authorizeObservation` for audit
// logging; every side-effecting operation goes through the approval queue.
//
// Approval model:
//   submitAction (requires approval): markRead, markUnread, moveToFolder, createReplyDraft,
//                                     createReplyAllDraft
//   authorizeObservation (audit-only): all reads
//
// Mutations return `void` and nothing is written to the mailbox before approval — no draft is
// created, no id is minted, and no later read reflects a queued action. Simulating pending state
// would mean carrying an overlay for a mailbox that other clients mutate concurrently; reporting
// nothing is honest about what has actually happened. `awaitDecision` is therefore set on every
// action so the agent stops rather than re-reading a mailbox its action has not reached yet.
//
// Every message id handled here is a Graph immutable id (see graph-api.ts). That is what lets an
// action queued now still apply after the message has moved folders in the meantime.

import { DurableObject, RpcStub, RpcTarget } from "cloudflare:workers";
import { validateRpc } from "capnweb-validate";
import {
  ActionKind, AgentCatalog, ApprovalQueue, Cursor, Gatekeeper,
  GatekeeperUserVerifier, ObservationAuthorizer, ResourceDescription, boundAgentCatalog,
} from "@gadgets/workshop-shared/gatekeeper";
import { AccessTokenCache, AccessTokenRequest } from "./auth-retry";
import { GraphMailApi, MAX_REPLY_BODY_BYTES, validateSearchQuery } from "./graph-api";
import type {
  OutlookAttachmentInfo, OutlookFolder, OutlookFolderEntry, OutlookFolderInfo, OutlookMailSession,
  OutlookMessage, OutlookMessageEntry, OutlookMessageInfo,
} from "./types";
import TYPES_CODE from "./types.txt";
import type { Env, UserAccount } from "./microsoft";

/** Canonical URL of the mailbox resource. The mailbox is a singleton per connected account. */
export const OUTLOOK_MAIL_URL = "https://outlook.office.com/mail/";

/** Ceiling on pages a single cursor will pull, so an agent cannot walk an entire mailbox. */
const MAX_CURSOR_PAGES = 40;

/** Ceiling on queued-but-undecided actions, matching the Gmail gatekeeper. */
const MAX_PENDING_ACTIONS = 100;

/**
 * Ceiling on folders offered to the agent catalog. Folders are the one item class here that grows
 * with the mailbox, so they are bounded well under the shared AGENT_CATALOG_MAX_ENTRIES ceiling;
 * folders past the cap stay reachable through the session's listFolders().
 */
const MAX_CATALOG_FOLDERS = 25;

/** Longest single-line value echoed into an approval title. */
const MAX_APPROVAL_TITLE_CHARS = 200;

class PendingActionStore<Action> {
  #kv: DurableObjectStorage["kv"];

  constructor(kv: DurableObjectStorage["kv"]) {
    this.#kv = kv;
  }

  #actionKey(id: number): string {
    return `pending:action:${id}`;
  }

  submit(action: Action): number {
    let id = this.#kv.get<number>("pending:nextActionId") ?? 1;
    this.#kv.put("pending:nextActionId", id + 1);
    this.#kv.put(this.#actionKey(id), action);
    return id;
  }

  get(id: number): Action | undefined {
    return this.#kv.get<Action>(this.#actionKey(id));
  }

  list(): {id: number, action: Action}[] {
    return [...this.#kv.list<Action>({prefix: "pending:action:"})]
        .map(([key, action]) => ({id: Number(key.slice("pending:action:".length)), action}))
        .filter(({id}) => Number.isFinite(id))
        .toSorted((a, b) => a.id - b.id);
  }

  remove(id: number): void {
    this.#kv.delete(this.#actionKey(id));
  }
}

// ── Action types ────────────────────────────────────────────────────
//
// Actions store the message's immutable id and the caller's intent, never a pre-built API payload:
// the executor rebuilds the request at approval time against the mailbox as it is then.

type OutlookMailAction =
  | { type: "setRead"; messageId: string; read: boolean }
  | { type: "move"; messageId: string; destinationFolderId: string }
  | { type: "replyDraft"; messageId: string; body: string; replyAll: boolean };

/** Grouped so a user can auto-approve read-state flips without also auto-approving moves/drafts. */
const MARK_READ_ACTION: ActionKind = { tag: "outlookMarkRead", label: "Mark messages read/unread" };
const MOVE_MESSAGE_ACTION: ActionKind = { tag: "outlookMoveMessage", label: "Move messages" };
const REPLY_DRAFT_ACTION: ActionKind = { tag: "outlookReplyDraft", label: "Create reply drafts" };

// ── Session context ─────────────────────────────────────────────────

type OutlookMailSessionContext = {
  api: GraphMailApi;
  approvalQueue: RpcStub<ApprovalQueue>;
  pendingActions: PendingActionStore<OutlookMailAction>;
};

function sanitizeApprovalTitle(value: string): string {
  return value.replace(/[\r\n]+/g, " ").slice(0, MAX_APPROVAL_TITLE_CHARS);
}

function formatApprovalField(label: string, value: string): string {
  // Use a fence longer than any backtick run in the value, so untrusted email fields render
  // verbatim and cannot forge surrounding approval Markdown.
  let fence = "```";
  while (value.includes(fence)) fence += "`";
  return `**${label}:**\n\n${fence}\n${value}\n${fence}`;
}

function describeMessage(info: OutlookMessageInfo): string {
  return [
    formatApprovalField("Subject", info.subject),
    formatApprovalField("From", info.from?.address ?? "(unknown sender)"),
    formatApprovalField("Received", info.receivedAt.toISOString()),
  ].join("\n\n");
}

/**
 * One line per attachment for an approval prompt. Sender-chosen filenames and MIME types go through
 * `formatApprovalField` at the call site, which is what keeps them from forging Markdown.
 */
function describeAttachments(attachments: OutlookAttachmentInfo[]): string {
  if (attachments.length === 0) return "(none)";
  return attachments.map(attachment =>
      `${attachment.name} — ${attachment.mimeType}, ${attachment.sizeBytes} bytes, ` +
      `${attachment.kind}${attachment.isInline ? ", inline" : ""}`).join("\n");
}

async function submitOutlookAction(
    ctx: OutlookMailSessionContext,
    action: OutlookMailAction,
    desc: { title: string; description: string; actionKind: ActionKind }): Promise<void> {
  if (ctx.pendingActions.list().length >= MAX_PENDING_ACTIONS) {
    throw new Error(
        "Too many pending Outlook actions. Resolve existing actions before adding more.");
  }
  let actionId = ctx.pendingActions.submit(action);
  try {
    await ctx.approvalQueue.submitAction(actionId, {
      ...desc,
      implementsRevert: false,
      // Nothing about this action is visible to later reads until it is applied, so an agent that
      // kept working would observe a mailbox where its action never happened.
      awaitDecision: true,
    });
  } catch (err) {
    ctx.pendingActions.remove(actionId);
    throw err;
  }
}

// ── OutlookMessageCursorImpl ────────────────────────────────────────
// Lazily pulls pages from Graph as the gadget calls next(), following `@odata.nextLink`. The cursor
// itself is a capability: listMessages()/search() authorizes its creation, and each next()
// separately authorizes the page it returns.

@validateRpc()
class OutlookMessageCursorImpl extends RpcTarget implements Cursor<OutlookMessageEntry> {
  #ctx: OutlookMailSessionContext;
  #firstPage: () => Promise<{ items: OutlookMessageInfo[]; nextLink?: string }>;
  #nextLink: string | undefined;
  #started = false;
  #exhausted = false;
  #pages = 0;
  #tail: Promise<void> = Promise.resolve();
  #describeScope: string;

  constructor(
      ctx: OutlookMailSessionContext,
      describeScope: string,
      firstPage: () => Promise<{ items: OutlookMessageInfo[]; nextLink?: string }>) {
    super();
    this.#ctx = ctx;
    this.#firstPage = firstPage;
    this.#describeScope = describeScope;
  }

  next(): Promise<OutlookMessageEntry[] | null> {
    // Serialized: two overlapping next() calls would both read `#nextLink` before either advanced
    // it, returning the same page twice.
    const result = this.#tail.then(() => this.#nextPage());
    this.#tail = result.then(() => undefined, () => undefined);
    return result;
  }

  async #nextPage(): Promise<OutlookMessageEntry[] | null> {
    if (this.#exhausted) return null;
    if (this.#pages >= MAX_CURSOR_PAGES) {
      throw new Error(
          `This cursor has already returned ${MAX_CURSOR_PAGES} pages. Narrow the search instead ` +
          "of paging further.");
    }

    // Pagination state is staged in locals and only written back once the observation has been
    // authorized. Committing first would let a refused (or failed) authorization advance the
    // cursor, so the retry the caller makes would silently return the page AFTER the one it never
    // received.
    let page: { items: OutlookMessageInfo[]; nextLink?: string };
    let nextLink = this.#nextLink;
    let started = this.#started;
    let pages = this.#pages;
    let skipped = 0;
    do {
      if (!started) {
        page = await this.#firstPage();
        started = true;
      } else if (nextLink) {
        page = await this.#ctx.api.nextMessagePage(nextLink);
      } else {
        this.#exhausted = true;
        return null;
      }
      nextLink = page.nextLink;
      pages++;
      // Graph can answer with an empty page that still carries a next link (e.g. every message on
      // the page was filtered out server-side). Skip those rather than reporting exhaustion.
      skipped++;
    } while (page.items.length === 0 && page.nextLink && skipped < 5 &&
             pages < MAX_CURSOR_PAGES);

    if (page.items.length === 0) {
      // Nothing was returned to the caller, so only the terminal state is recorded — an empty page
      // is not a page anyone can be asked to re-read.
      this.#started = started;
      this.#nextLink = nextLink;
      this.#pages = pages;
      this.#exhausted = !page.nextLink;
      return null;
    }

    const entries = page.items.map(info => ({
      info,
      message: new OutlookMessageStub(this.#ctx, info.id, info) as OutlookMessage,
    }));

    await this.#ctx.approvalQueue.authorizeObservation({
      title: `Read ${entries.length} Outlook messages`,
      description:
          `Fetch the next page of messages from ${this.#describeScope}.\n\n` +
          formatApprovalField("Subjects", entries.map(entry => entry.info.subject).join("\n")),
    });

    this.#started = started;
    this.#nextLink = nextLink;
    this.#pages = pages;
    if (!page.nextLink) this.#exhausted = true;
    return entries;
  }
}

// ── OutlookFolderStub ───────────────────────────────────────────────

@validateRpc()
class OutlookFolderStub extends RpcTarget implements OutlookFolder {
  #ctx: OutlookMailSessionContext;
  #folderId: string;
  #cachedInfo: OutlookFolderInfo | undefined;

  constructor(ctx: OutlookMailSessionContext, folderId: string, cachedInfo?: OutlookFolderInfo) {
    super();
    this.#ctx = ctx;
    this.#folderId = folderId;
    this.#cachedInfo = cachedInfo;
  }

  async getInfo(): Promise<OutlookFolderInfo> {
    let info = this.#cachedInfo ?? await this.#ctx.api.getFolder(this.#folderId);
    this.#cachedInfo = info;

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Outlook folder: ${info.name}`),
      description:
          "Read metadata for this mail folder.\n\n" +
          formatApprovalField("Folder", info.name),
    });

    return info;
  }

  async listMessages(): Promise<Cursor<OutlookMessageEntry>> {
    let info = this.#cachedInfo;
    let scope = info ? `the "${info.name}" folder` : "the selected mail folder";

    await this.#ctx.approvalQueue.authorizeObservation({
      title: "List Outlook messages in a folder",
      description: `Create a cursor over the most recent messages in ${scope}.`,
    });

    return new OutlookMessageCursorImpl(this.#ctx, scope,
        () => this.#ctx.api.listMessages({ folderId: this.#folderId }));
  }
}

// ── OutlookMessageStub ──────────────────────────────────────────────

@validateRpc()
class OutlookMessageStub extends RpcTarget implements OutlookMessage {
  #ctx: OutlookMailSessionContext;
  #messageId: string;
  #cachedInfo: OutlookMessageInfo | undefined;

  constructor(ctx: OutlookMailSessionContext, messageId: string, cachedInfo?: OutlookMessageInfo) {
    super();
    this.#ctx = ctx;
    this.#messageId = messageId;
    this.#cachedInfo = cachedInfo;
  }

  async #ensureInfo(): Promise<OutlookMessageInfo> {
    if (!this.#cachedInfo) {
      this.#cachedInfo = await this.#ctx.api.getMessage(this.#messageId);
    }
    return this.#cachedInfo;
  }

  async getMetadata(): Promise<OutlookMessageInfo> {
    // Always re-read: metadata carries the read state and current folder, which other mail clients
    // change constantly, and a stale answer would be indistinguishable from a fresh one.
    let info = await this.#ctx.api.getMessage(this.#messageId);
    this.#cachedInfo = info;

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Message info: ${info.subject}`),
      description: `Read metadata for an Outlook message.\n\n${describeMessage(info)}`,
    });

    return info;
  }

  async getBody(): Promise<string> {
    let info = await this.#ensureInfo();
    let body = await this.#ctx.api.getMessageBody(this.#messageId);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Read message: ${info.subject}`),
      description: `Read the body of an Outlook message.\n\n${describeMessage(info)}`,
    });

    return body;
  }

  /** Attachment metadata for this message. Reads no attachment content. */
  async listAttachments(): Promise<OutlookAttachmentInfo[]> {
    let info = await this.#ensureInfo();
    let attachments = await this.#ctx.api.listAttachments(this.#messageId);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`List ${attachments.length} attachments: ${info.subject}`),
      description:
          "List the attachments on an Outlook message.\n\n" +
          `${describeMessage(info)}\n\n` +
          formatApprovalField("Attachments", describeAttachments(attachments)),
    });

    return attachments;
  }

  /**
   * One attachment's bytes. The size cap, and the refusal of kinds that carry no bytes, live in the
   * Graph client, so nothing oversized is downloaded to be rejected here.
   */
  async getAttachmentContent(attachmentId: string): Promise<ArrayBuffer> {
    if (!attachmentId) throw new Error("getAttachmentContent() requires an attachment id.");
    let info = await this.#ensureInfo();
    // Like getBody(), the fetch may precede the authorization; what matters is that nothing is
    // returned to the caller until the observation has been authorized.
    let { info: attachment, content } =
        await this.#ctx.api.getAttachmentBytes(this.#messageId, attachmentId);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(
          `Read attachment: ${attachment.name} (${content.byteLength} bytes) — ${info.subject}`),
      description:
          "Read the contents of an attachment on an Outlook message.\n\n" +
          `${describeMessage(info)}\n\n` +
          formatApprovalField("Attachment", describeAttachments([attachment])),
    });

    return content;
  }

  /**
   * Queue "mark as read". Returns as soon as the action is queued; the mailbox is unchanged until
   * the action is approved, and this session never reports the outcome.
   */
  async markRead(): Promise<void> {
    await this.#submitReadState(true);
  }

  /**
   * Queue "mark as unread". Same queued-for-approval semantics as `markRead()`.
   */
  async markUnread(): Promise<void> {
    await this.#submitReadState(false);
  }

  async #submitReadState(read: boolean): Promise<void> {
    let info = await this.#readInfoForAction(read ? "mark read" : "mark unread");
    await submitOutlookAction(
        this.#ctx,
        { type: "setRead", messageId: this.#messageId, read },
        {
          title: sanitizeApprovalTitle(
              `${read ? "Mark read" : "Mark unread"}: ${info.subject}`),
          description:
              `Mark this Outlook message as ${read ? "read" : "unread"}.\n\n` +
              describeMessage(info),
          actionKind: MARK_READ_ACTION,
        });
  }

  /**
   * Queue a move to `folderId`. Returns as soon as the action is queued; the message does not move
   * until the action is approved, and no result is reported back.
   */
  async moveToFolder(folderId: string): Promise<void> {
    if (!folderId) throw new Error("moveToFolder() requires a folder id.");
    let info = await this.#readInfoForAction("move");
    // Reading the destination is what turns an opaque id into something a human can approve.
    let folder = await this.#ctx.api.getFolder(folderId);
    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Read destination folder: ${folder.name}`),
      description: "Read the destination folder's name to describe a pending move.",
    });

    await submitOutlookAction(
        this.#ctx,
        { type: "move", messageId: this.#messageId, destinationFolderId: folder.id },
        {
          title: sanitizeApprovalTitle(`Move to ${folder.name}: ${info.subject}`),
          description:
              "Move this Outlook message to another folder.\n\n" +
              `${describeMessage(info)}\n\n` +
              formatApprovalField("Destination folder", folder.name),
          actionKind: MOVE_MESSAGE_ACTION,
        });
  }

  /**
   * Queue a draft reply to the sender. Returns as soon as the action is queued; no draft exists
   * until the action is approved, nothing is ever sent, and no draft id is returned.
   */
  async createReplyDraft(body: string): Promise<void> {
    await this.#submitReplyDraft(body, false);
  }

  /**
   * Queue a draft reply to every recipient. Same queued-for-approval semantics as
   * `createReplyDraft()`.
   */
  async createReplyAllDraft(body: string): Promise<void> {
    await this.#submitReplyDraft(body, true);
  }

  async #submitReplyDraft(body: string, replyAll: boolean): Promise<void> {
    if (new TextEncoder().encode(body).byteLength > MAX_REPLY_BODY_BYTES) {
      throw new Error(`Reply body must be at most ${MAX_REPLY_BODY_BYTES} bytes.`);
    }
    let info = await this.#readInfoForAction("draft a reply to");
    let recipients = replyAll
        ? [info.from?.address, ...info.to.map(entry => entry.address),
           ...info.cc.map(entry => entry.address)].filter(Boolean).join(", ")
        : info.from?.address ?? "(unknown sender)";

    await submitOutlookAction(
        this.#ctx,
        { type: "replyDraft", messageId: this.#messageId, body, replyAll },
        {
          title: sanitizeApprovalTitle(
              `${replyAll ? "Reply-all draft" : "Reply draft"}: ${info.subject}`),
          description:
              "Create a reply draft in the Outlook mailbox. The draft is not sent; it is left in " +
              "Drafts for the user to review.\n\n" +
              `${describeMessage(info)}\n\n` +
              `${formatApprovalField("Draft recipients", recipients)}\n\n` +
              formatApprovalField("Draft body", body),
          actionKind: REPLY_DRAFT_ACTION,
        });
  }

  /** Read the message so the approval prompt can describe what is being acted on. */
  async #readInfoForAction(what: string): Promise<OutlookMessageInfo> {
    let info = await this.#ensureInfo();
    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Read message before ${what}: ${info.subject}`),
      description: `Read the message details needed to describe this pending action.\n\n` +
          describeMessage(info),
    });
    return info;
  }
}

// ── OutlookMailSessionImpl ──────────────────────────────────────────

@validateRpc()
class OutlookMailSessionImpl extends RpcTarget implements OutlookMailSession {
  #ctx: OutlookMailSessionContext;

  constructor(ctx: OutlookMailSessionContext) {
    super();
    this.#ctx = ctx;
  }

  async listMessages(): Promise<Cursor<OutlookMessageEntry>> {
    await this.#ctx.approvalQueue.authorizeObservation({
      title: "List Outlook messages",
      description: "Create a cursor over the most recent messages in the connected mailbox.",
    });
    return new OutlookMessageCursorImpl(this.#ctx, "the connected mailbox",
        () => this.#ctx.api.listMessages());
  }

  async search(query: string): Promise<Cursor<OutlookMessageEntry>> {
    // Validated before the approval prompt, so a malformed query fails immediately instead of
    // asking a human to approve a search that cannot run.
    let validated = validateSearchQuery(query);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: "Search Outlook",
      description:
          "Create a cursor over mailbox messages matching this search.\n\n" +
          formatApprovalField("Query", validated),
    });
    return new OutlookMessageCursorImpl(this.#ctx, "this mailbox search",
        () => this.#ctx.api.searchMessages(validated));
  }

  /**
   * Re-derive a message capability from an id the agent already holds.
   *
   * No observation is authorized here, and none is owed: minting the stub reads nothing from the
   * mailbox, and every read made through it authorizes itself individually. An id that names no
   * message fails on the first call, not here.
   */
  async getMessage(messageId: string): Promise<OutlookMessage> {
    if (!messageId) throw new Error("getMessage() requires a message id.");
    return new OutlookMessageStub(this.#ctx, messageId) as OutlookMessage;
  }

  async listFolders(): Promise<OutlookFolderEntry[]> {
    let folders = await this.#ctx.api.listFolders();

    await this.#ctx.approvalQueue.authorizeObservation({
      title: `List ${folders.length} Outlook folders`,
      description:
          "List the mail folders in the connected mailbox.\n\n" +
          formatApprovalField("Folders", folders.map(folder => folder.name).join("\n")),
    });

    return folders.map(info => ({
      info,
      folder: new OutlookFolderStub(this.#ctx, info.id, info) as OutlookFolder,
    }));
  }
}

// =======================================================================================

export type OutlookMailGatekeeperImplProps = {
  userObjectId: string;
}

@validateRpc()
export class OutlookMailGatekeeperImpl
    extends DurableObject<Env, OutlookMailGatekeeperImplProps>
    implements Gatekeeper<OutlookMailSession> {
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
   * account that refreshes happily while every mailbox call fails.
   *
   * This object's own token memo is dropped first. It is the only copy the account authority cannot
   * reach, and it outlives the reconnect that fixes the account: keeping it would serve the rejected
   * token again after the user has reconnected, re-reporting a death that no longer exists.
   */
  #api(): GraphMailApi {
    return new GraphMailApi(opts => this.#getAccessToken(opts), {
      onCredentialsRejected: async (detail: string) => {
        this.#tokens.invalidate();
        await this.#account().reportCredentialsRejected(detail);
      },
    });
  }

  async describe(): Promise<ResourceDescription> {
    return {
      url: OUTLOOK_MAIL_URL,
      title: "Outlook Mailbox",
      snippet: "Your Microsoft 365 mailbox",
      suggestedBindingName: "OUTLOOK_MAILBOX",
      tsType: "OutlookMailSession",
    };
  }

  async getTypeScriptTypes(): Promise<string> {
    return TYPES_CODE;
  }

  async getAutoApprovableActions(): Promise<ActionKind[]> {
    // Nothing is auto-approvable: every action writes to a mailbox the user owns, and the read/
    // unread flip is the only reversible one — not enough to make it safe by default.
    return [];
  }

  async startSession(approvalQueue: RpcStub<ApprovalQueue>): Promise<OutlookMailSession> {
    return new OutlookMailSessionImpl({
      api: this.#api(),
      approvalQueue: approvalQueue.dup(),
      pendingActions: new PendingActionStore<OutlookMailAction>(this.ctx.storage.kv),
    });
  }

  /**
   * Folder names, so an agent can discover where mail lives (and where it could be moved) without
   * paging the session API.
   */
  async getAgentCatalog(
      authorizer: RpcStub<ObservationAuthorizer>): Promise<AgentCatalog | null> {
    let folders = await this.#api().listFolders();
    let entries = folders
        .slice(0, MAX_CATALOG_FOLDERS)
        .map(folder => ({
          id: folder.id,
          title: folder.name,
          description:
              `Outlook mail folder with ${folder.totalItemCount} message(s), ` +
              `${folder.unreadItemCount} unread.`,
        }));
    let catalog = boundAgentCatalog(entries);
    // A mailbox can hold more folders than the catalog advertises; boundAgentCatalog only flags its
    // own shared ceiling, so the drop at MAX_CATALOG_FOLDERS is reported here.
    if (folders.length > MAX_CATALOG_FOLDERS) catalog.truncated = true;
    if (catalog.entries.length > 0) {
      await authorizer.authorizeObservation({
        title: "Outlook folder catalog",
        description: `Listed ${catalog.entries.length} Outlook mail folder(s).`,
      });
    }
    return catalog;
  }

  // ---------------------------------------------------------------------------
  // The only place this gatekeeper writes to the mailbox. Everything the session offers either
  // reads or queues an action that lands here after approval.

  async applyAction(actionId: number): Promise<void> {
    const pendingActions = new PendingActionStore<OutlookMailAction>(this.ctx.storage.kv);
    const action = pendingActions.get(actionId);
    if (!action) throw new Error(`Unknown pending Outlook action: ${actionId}`);

    const api = this.#api();
    switch (action.type) {
      case "setRead":
        await api.setMessageRead(action.messageId, action.read);
        break;
      case "move":
        await api.moveMessage(action.messageId, action.destinationFolderId);
        break;
      case "replyDraft":
        await api.createReplyDraft(action.messageId, action.body, action.replyAll);
        break;
      default:
        action satisfies never;
        throw new Error(`unknown action type: ${(action as {type: string}).type}`);
    }

    pendingActions.remove(actionId);
  }

  async rejectAction(actionId: number): Promise<void | {restart?: boolean}> {
    const pendingActions = new PendingActionStore<OutlookMailAction>(this.ctx.storage.kv);
    if (!pendingActions.get(actionId)) {
      throw new Error(`Unknown pending Outlook action: ${actionId}`);
    }
    pendingActions.remove(actionId);
  }

  async revertAction(_action: number):
      Promise<void | {message?: string, canRetry?: boolean, restart?: boolean}> {
    // A move cannot be undone without recording where the message came from, and a created draft
    // may already have been edited or sent by the user, so nothing here is safely reversible.
    // Actions are submitted with implementsRevert: false, so the UI never offers this.
    throw new Error("revert is not implemented");
  }

  /**
   * Observer tracking — private-only. A mailbox has no per-recipient ACL to verify an observer
   * against, and full access to someone's mail is too personal to extend to a non-owner, so
   * addObserver always throws. removeObserver is a no-op since none is ever recorded.
   */
  async addObserver(_id: string, _user: Fetcher<GatekeeperUserVerifier>): Promise<void> {
    throw new Error(
        "Outlook mailbox data cannot be shared with other users: this workspace reads a personal " +
        "mailbox, which may only be observed by its owner.");
  }

  async removeObserver(_id: string): Promise<void> {}
}
