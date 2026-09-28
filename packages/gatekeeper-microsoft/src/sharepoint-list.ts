// SharePoint list gatekeeper.
//
// One binding is one list, fixed at introduction: the site and list ids live in the Durable Object's
// props and nothing in the session can point at another list. Reads go through
// `authorizeObservation` for audit logging; creating an item goes through the approval queue.
//
// Approval model:
//   submitAction (auto-approvable): createItem
//   authorizeObservation (audit-only): getColumns, every getItems page, getItem
//
// Unlike the mailbox next door, the create action is auto-approvable, so the common case is that the
// workspace has enabled the "Create list items" rule and the overseer applies the action moments
// after it is submitted. `createItem` therefore waits a bounded time for its own action to be
// applied and returns the real item id when that happens. When it does not — the drain applies one
// action at a time per gatekeeper, and the rule may simply be off — it returns a provisional
// `pending-<actionId>` id, and `getItem` resolves that id later against a durable outcome record.
// That record is what makes every ending (applied, failed, throttled, rejected, or blocked behind an
// earlier failure) visible to whoever submitted, long after the waiting call gave up and long after
// this object was evicted.
//
// `applyAction` is idempotent: it short-circuits on an outcome that says the item already exists, so
// an overseer retry after a crash in its own bookkeeping cannot create the item twice. The one
// window left is a crash between SharePoint's 201 and the transaction that records it, which would
// need an idempotency column in the user's own list to close; that is not a schema this gatekeeper
// imposes.

import { DurableObject, RpcStub, RpcTarget } from "cloudflare:workers";
import { validateRpc } from "capnweb-validate";
import {
  ActionKind, ApprovalQueue, Cursor, Gatekeeper, GatekeeperUserVerifier, ResourceDescription,
} from "@gadgets/workshop-shared/gatekeeper";
import { buildDescription, type RenderedDescription } from "@gadgets/gatekeeper-kit/action-description";
import { formatApprovalField, sanitizeApprovalTitle } from "./approval-text";
import { AccessTokenCache, AccessTokenRequest, RetryAfterPolicy } from "./auth-retry";
import { GraphApiError, truncate } from "./graph-api";
import {
  ColumnDefinition, GraphSharePointApi, ListItem, buildItemsFilter, sharePointThrottled,
} from "./graph-sharepoint-api";
import { MAX_PENDING_ACTIONS, PendingActionStore } from "./pending-actions";
import type {
  CreateItemResult, GetItemsOptions, SharePointListSession,
} from "./sharepoint-types";
import SHAREPOINT_TYPES_CODE from "./sharepoint-types.txt";
import type { Env, SharePointVerifierApi, UserAccount } from "./microsoft";

/** Ceiling on pages a single cursor will pull, so an agent cannot walk an entire list. */
const MAX_CURSOR_PAGES = 40;

/** Longest caller- or provider-supplied name echoed back in an error message. */
const MAX_NAME_ECHO_CHARS = 100;

/** Longest failure text kept in an outcome record, where it is read back days later. */
const MAX_OUTCOME_REASON_CHARS = 300;

/**
 * How long `createItem` waits for the auto-approval drain to apply its own action before answering
 * with a provisional id.
 *
 * Long enough to cover a Graph write plus the overseer's round trip, short enough that an agent is
 * not blocked when the rule is off and nobody is at the keyboard. Overridable per instance (the
 * constructor's third parameter) so tests can collapse it; `0` disables the wait entirely, which is
 * the documented fallback if the platform's delivery semantics make waiting pointless.
 */
const WAIT_FOR_APPLY_MS = 5000;

/** Throttling ceiling for reads: an agent is waiting on these synchronously. */
const READ_RETRY_AFTER: RetryAfterPolicy = { maxWaitMs: 10_000, tooLong: sharePointThrottled };

/**
 * Throttling ceiling for applies. Six times the read ceiling: an approval queue is behind this call
 * and nobody is watching it, so sitting out a real SharePoint throttle is worth more than failing
 * fast and leaving the action stuck at the head of the drain.
 */
const APPLY_RETRY_AFTER: RetryAfterPolicy = { maxWaitMs: 60_000, tooLong: sharePointThrottled };

/** Where the list's display name and link are cached, so `describe()` survives a Graph blip. */
const META_KEY = "sp:meta";

/** Prefix of the durable record describing how one submitted action ended. */
const OUTCOME_PREFIX = "sp:outcome:";

/** Prefix of the ids `createItem` returns for an action that has not been applied yet. */
const PROVISIONAL_PREFIX = "pending-";

/** A SharePoint list item id: a positive whole number, as a string. */
const ITEM_ID_PATTERN = /^[1-9][0-9]*$/;

/** The one action this gatekeeper submits. Auto-approvable, per introduced list in the workspace. */
const CREATE_ITEM_ACTION: ActionKind = { tag: "create-item", label: "Create list items" };

// ── Action and outcome records ──────────────────────────────────────

type SharePointListAction = {
  type: "createItem";
  /** Validated against the list's schema before submission, so apply time rarely sees a 4xx. */
  fields: Record<string, unknown>;
};

/** The list's name and link, as `describe()` last read them. */
type ListMeta = {
  displayName: string;
  webUrl: string;
};

/**
 * How a submitted action ended. Durable, because the submitter usually learns the answer from a
 * later `getItem()` call rather than from the call that submitted, and the object that applied the
 * action may be long evicted by then.
 *
 * A `failed` action stays pending on purpose: the platform's drain stops rather than skipping ahead,
 * so a human has to decide whether to retry or reject it, and this record is what tells them (and
 * every other submitter) why everything behind it is stuck.
 */
type ActionOutcome =
  | { state: "applied"; itemId: string; webUrl?: string }
  | { state: "failed"; reason: string; throttled: boolean }
  | { state: "rejected" };

/** Storage key of an action's outcome record. */
function outcomeKey(actionId: number): string {
  return `${OUTCOME_PREFIX}${actionId}`;
}

/** The id `createItem` hands back for an action that has not been applied yet. */
function provisionalItemId(actionId: number): string {
  return `${PROVISIONAL_PREFIX}${actionId}`;
}

/** The action id inside a provisional item id, or null when the id is not one. */
function parseProvisionalItemId(id: string): number | null {
  if (!id.startsWith(PROVISIONAL_PREFIX)) return null;
  let rest = id.slice(PROVISIONAL_PREFIX.length);
  if (!ITEM_ID_PATTERN.test(rest)) return null;
  return Number(rest);
}

/**
 * A submission that will not become an item.
 *
 * Distinguishable from a transport or validation failure so a form can show the user what happened
 * to their submission rather than a generic error.
 */
export class SharePointSubmissionError extends Error {
  readonly outcome: "failed" | "rejected";
  readonly throttled: boolean;

  constructor(message: string, outcome: "failed" | "rejected", throttled = false) {
    super(message);
    this.name = "SharePointSubmissionError";
    this.outcome = outcome;
    this.throttled = throttled;
  }
}

/** The error a decided-against submission is reported with, wherever it is discovered. */
function submissionError(actionId: number, outcome: ActionOutcome): SharePointSubmissionError {
  if (outcome.state === "rejected") {
    return new SharePointSubmissionError(
        `Submission ${actionId} was rejected by an approver.`, "rejected");
  }
  if (outcome.state === "failed") {
    return new SharePointSubmissionError(
        outcome.throttled
            ? `Submission ${actionId} failed: SharePoint throttled the write. It is still queued, ` +
              "so an approver can retry it from the workspace's activity list."
            : `Submission ${actionId} failed: ${outcome.reason}`,
        "failed", outcome.throttled);
  }
  // Only reached if a caller asks for the error form of an applied action, which is a bug here
  // rather than anything the submitter did.
  throw new Error(`Submission ${actionId} did not fail.`);
}

/** The outcome recorded when a Graph write throws. Reason is sanitised: it is shown to a user. */
function failedOutcome(error: unknown): ActionOutcome {
  // `sharePointThrottled` builds a 429 too, so one status covers both "the server throttled us" and
  // "the server asked for a longer wait than the apply policy will sit out".
  let throttled = error instanceof GraphApiError && error.status === 429;
  let reason = error instanceof Error && error.message ? error.message : String(error);
  return {
    state: "failed",
    reason: truncate(reason.replace(/[\r\n]+/g, " "), MAX_OUTCOME_REASON_CHARS),
    throttled,
  };
}

// ── Approval text ───────────────────────────────────────────────────

/** Echo a caller- or provider-supplied name inside an error, without letting it carry structure. */
function echoName(value: unknown): string {
  return truncate(String(value).replace(/[\r\n]+/g, " "), MAX_NAME_ECHO_CHARS);
}

function displayValue(value: unknown): string {
  if (value === null) return "(cleared)";
  return typeof value === "string" ? value : String(value);
}

/**
 * What an approver reads before allowing an item to be created: one field per column value the
 * create sends. Labels are internal column names, not display names: an internal name is
 * `[A-Za-z0-9_]+` by construction, while a display name is renameable by anyone who can edit the
 * list. Values go in builder fields, which approval surfaces show literally, so a submitted
 * `<script>` or a run of backticks cannot restyle the prompt around it.
 */
function describeCreateItem(fields: Record<string, unknown>): RenderedDescription {
  let builder = buildDescription("Create a new item in this SharePoint list with the values below.");
  let entries = Object.entries(fields);
  if (entries.length === 0) builder.prose("_This item sets no column values._");
  for (let [name, value] of entries) builder.inline(name, displayValue(value));
  return builder.finish();
}

// ── Field validation ────────────────────────────────────────────────

/**
 * Check a create against the list's real schema, before anything is queued.
 *
 * Everything an approver would otherwise wait hours to see fail is caught here: a column that does
 * not exist, a value of the wrong type, a choice the column does not offer, a required column left
 * out, and the column kinds this gatekeeper deliberately cannot write. That matters more than usual
 * because a failed apply stops the drain: one bad submission blocks every later one on the same
 * list until a human clears it.
 *
 * Returns the values as they will be sent to Graph — dates normalised to ISO, everything else
 * unchanged — in a fresh object, so a later mutation of the caller's argument cannot change what was
 * approved.
 */
export function validateFields(
    fields: Record<string, unknown>, columns: ColumnDefinition[]): Record<string, unknown> {
  if (typeof fields !== "object" || fields === null || Array.isArray(fields)) {
    throw new Error("createItem() takes an object of column values, keyed by internal column name.");
  }

  let validated: [string, unknown][] = [];
  for (let [name, value] of Object.entries(fields)) {
    // An absent value and an explicitly undefined one say the same thing, so the column is simply
    // left out and the required check below decides whether that is allowed.
    if (value === undefined) continue;

    let column = columns.find(candidate => candidate.name === name);
    if (!column) {
      throw new Error(
          `Unknown column "${echoName(name)}". Use internal column names from getColumns().`);
    }
    // `readOnly` is deliberately not consulted. The Graph client already drops every read-only
    // column except `Title`, which some lists mark read-only in a view definition while still
    // accepting it on create — so the flag never reaches here meaning "unwritable", and honouring it
    // would refuse the one column every list has.
    validated.push([column.name, validateValue(column, value)]);
  }

  let provided = new Set(validated
      .filter(([, value]) => value !== null)
      .map(([name]) => name));
  for (let column of columns) {
    if (!column.required) continue;
    if (provided.has(column.name)) continue;
    throw new Error(
        `Column "${column.name}" is required by this list, so createItem() cannot leave it out.`);
  }

  return Object.fromEntries(validated);
}

/** One field value, checked against its column. Returns the value in the shape Graph expects. */
function validateValue(column: ColumnDefinition, value: unknown): unknown {
  if (column.type === "person" || column.type === "lookup") {
    throw new Error(
        `Column "${column.name}" is a ${column.type} column. SharePoint stores those as numeric ids ` +
        "of directory or list entries, which this connection cannot look up, so they can be read " +
        "but not written.");
  }
  if (column.type === "unsupported") {
    throw new Error(
        `Column "${column.name}" has a type this connection can read but not write (currency, ` +
        "calculated, hyperlink, managed metadata and similar columns).");
  }

  // `null` clears a column. Required columns cannot be cleared, which the required check catches:
  // it counts a null as "not provided".
  if (value === null) return null;

  switch (column.type) {
    case "text":
      if (typeof value !== "string") {
        throw new Error(`Column "${column.name}" expects text, but got ${describeType(value)}.`);
      }
      return value;
    case "number":
      if (typeof value !== "number" || !Number.isFinite(value)) {
        throw new Error(
            `Column "${column.name}" expects a finite number, but got ${describeType(value)}.`);
      }
      return value;
    case "boolean":
      if (typeof value !== "boolean") {
        throw new Error(
            `Column "${column.name}" expects true or false, but got ${describeType(value)}.`);
      }
      return value;
    case "dateTime":
      return validateDateTime(column, value);
    case "choice": {
      if (typeof value !== "string") {
        throw new Error(
            `Column "${column.name}" expects one of its choices, but got ${describeType(value)}.`);
      }
      // A choice column whose options Graph did not report has nothing to check against; refusing
      // every value there would make the column unusable rather than safe.
      let choices = column.choices ?? [];
      if (choices.length > 0 && !choices.includes(value)) {
        throw new Error(
            `Column "${column.name}" does not offer the choice "${echoName(value)}". Allowed: ` +
            `${choices.join(", ")}.`);
      }
      return value;
    }
    default:
      column.type satisfies never;
      throw new Error(`Column "${column.name}" has an unrecognized type.`);
  }
}

function validateDateTime(column: ColumnDefinition, value: unknown): string {
  // A Date survives the RPC boundary, and a form is far likelier to send a string; both end up as
  // the ISO timestamp Graph documents, so neither the caller's formatting nor its time zone
  // shorthand reaches SharePoint.
  let date = value instanceof Date ? value : typeof value === "string" ? new Date(value) : null;
  if (!date || Number.isNaN(date.valueOf())) {
    throw new Error(
        `Column "${column.name}" expects a date, e.g. "2026-09-15" or an ISO timestamp, but got ` +
        `${describeType(value)}.`);
  }
  return date.toISOString();
}

function describeType(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "a list";
  if (value instanceof Date) return "a date";
  switch (typeof value) {
    case "string": return "text";
    case "number": return "a number";
    case "boolean": return "true/false";
    case "object": return "an object";
    default: return typeof value;
  }
}

// ── Session context ─────────────────────────────────────────────────

/** A registered wait for one action's outcome, and the means to give it up. */
type OutcomeWait = {
  /** Resolves with the outcome, or null when the wait ran out (or was cancelled). */
  outcome: Promise<ActionOutcome | null>;
  cancel(): void;
};

type SharePointListSessionContext = {
  api: GraphSharePointApi;
  approvalQueue: RpcStub<ApprovalQueue>;
  pendingActions: PendingActionStore<SharePointListAction>;
  kv: DurableObjectStorage["kv"];
  siteId: string;
  listId: string;
  /** The list's display name, as `describe()` last saw it. Provider text: sanitize before echoing. */
  listName(): string;
  /** The list's columns, read once per Durable Object lifetime unless a create invalidates them. */
  columns(): Promise<ColumnDefinition[]>;
  /** Registers a waiter for `actionId` before the action is submitted. */
  waitForOutcome(actionId: number): OutcomeWait;
};

/** Is an older submission sitting failed at the head of the drain, blocking everything behind it? */
function blockedByOlderFailure(ctx: SharePointListSessionContext, actionId: number): boolean {
  return ctx.pendingActions.list().some(({ id }) =>
      id < actionId && ctx.kv.get<ActionOutcome>(outcomeKey(id))?.state === "failed");
}

// ── SharePointItemCursorImpl ────────────────────────────────────────
// Lazily pulls pages from Graph as the gadget calls next(), following `@odata.nextLink`. The cursor
// is a capability: getItems() authorizes its creation, and each next() separately authorizes the
// page it returns.

@validateRpc()
class SharePointItemCursorImpl extends RpcTarget implements Cursor<ListItem> {
  #ctx: SharePointListSessionContext;
  #firstPage: () => Promise<{ items: ListItem[]; nextLink?: string }>;
  #nextLink: string | undefined;
  #started = false;
  #exhausted = false;
  #pages = 0;
  #tail: Promise<void> = Promise.resolve();

  constructor(
      ctx: SharePointListSessionContext,
      firstPage: () => Promise<{ items: ListItem[]; nextLink?: string }>) {
    super();
    this.#ctx = ctx;
    this.#firstPage = firstPage;
  }

  next(): Promise<ListItem[] | null> {
    // Serialized: two overlapping next() calls would both read `#nextLink` before either advanced
    // it, returning the same page twice.
    const result = this.#tail.then(() => this.#nextPage());
    this.#tail = result.then(() => undefined, () => undefined);
    return result;
  }

  async #nextPage(): Promise<ListItem[] | null> {
    if (this.#exhausted) return null;
    if (this.#pages >= MAX_CURSOR_PAGES) {
      throw new Error(
          `This cursor has already returned ${MAX_CURSOR_PAGES} pages. Narrow the query with ` +
          "`where` instead of paging further.");
    }

    // Pagination state is staged in locals and only written back once the observation has been
    // authorized. Committing first would let a refused (or failed) authorization advance the cursor,
    // so the retry the caller makes would silently return the page AFTER the one it never received.
    let page: { items: ListItem[]; nextLink?: string };
    let nextLink = this.#nextLink;
    let started = this.#started;
    let pages = this.#pages;
    let skipped = 0;
    do {
      if (!started) {
        page = await this.#firstPage();
        started = true;
      } else if (nextLink) {
        page = await this.#ctx.api.nextItemPage(nextLink);
      } else {
        this.#exhausted = true;
        return null;
      }
      nextLink = page.nextLink;
      pages++;
      // A filtered page can come back empty while still carrying a next link. Skip those rather than
      // reporting exhaustion.
      skipped++;
    } while (page.items.length === 0 && page.nextLink && skipped < 5 && pages < MAX_CURSOR_PAGES);

    if (page.items.length === 0) {
      // Nothing was returned to the caller, so only the terminal state is recorded — an empty page
      // is not a page anyone can be asked to re-read.
      this.#started = started;
      this.#nextLink = nextLink;
      this.#pages = pages;
      this.#exhausted = !page.nextLink;
      return null;
    }

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(
          `Read ${page.items.length} items from ${this.#ctx.listName()}`),
      description:
          "Fetch the next page of items from this SharePoint list.\n\n" +
          formatApprovalField("Item ids", page.items.map(item => item.id).join(", ")),
    });

    this.#started = started;
    this.#nextLink = nextLink;
    this.#pages = pages;
    if (!page.nextLink) this.#exhausted = true;
    return page.items;
  }
}

// ── SharePointListSessionImpl ───────────────────────────────────────

@validateRpc()
class SharePointListSessionImpl extends RpcTarget implements SharePointListSession {
  #ctx: SharePointListSessionContext;

  constructor(ctx: SharePointListSessionContext) {
    super();
    this.#ctx = ctx;
  }

  async getColumns(): Promise<ColumnDefinition[]> {
    let columns = await this.#ctx.columns();

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Read column schema of ${this.#ctx.listName()}`),
      description:
          "Read the column names, types and choices of this SharePoint list.\n\n" +
          formatApprovalField("Columns", columns.map(column => column.name).join(", ")),
    });

    return columns;
  }

  async getItems(options?: GetItemsOptions): Promise<Cursor<ListItem>> {
    let columns = await this.#ctx.columns();
    let select = options?.select;
    let where = options?.where ?? [];

    // Validated before the approval prompt, so a bad column or an unfilterable type fails here
    // rather than on the first next() — nobody should be asked to authorize a query that cannot run.
    if (select) {
      for (let name of select) {
        if (!columns.some(column => column.name === name)) {
          throw new Error(
              `Unknown column "${echoName(name)}". Use internal column names from getColumns().`);
        }
      }
    }
    buildItemsFilter(where, columns);

    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`List items in ${this.#ctx.listName()}`),
      description:
          "Create a cursor over the items of this SharePoint list.\n\n" +
          formatApprovalField("Filter", where.length > 0
              ? where.map(clause => `${clause.column} ${clause.op} ${displayValue(clause.value)}`)
                  .join("\n")
              : "(none)"),
    });

    return new SharePointItemCursorImpl(this.#ctx, () =>
        this.#ctx.api.listItems(this.#ctx.siteId, this.#ctx.listId, {
          columns,
          ...(select ? { select } : {}),
          ...(where.length > 0 ? { where } : {}),
          ...(options?.top !== undefined ? { top: options.top } : {}),
        }));
  }

  async getItem(id: string): Promise<ListItem> {
    if (typeof id !== "string" || id === "") throw new Error("getItem() requires an item id.");

    // A provisional id is this gatekeeper's own bookkeeping, not a SharePoint id, so it is
    // recognized before the id is checked against SharePoint's format.
    let actionId = parseProvisionalItemId(id);
    if (actionId === null && !ITEM_ID_PATTERN.test(id)) {
      throw new Error(
          `"${echoName(id)}" is not a SharePoint item id. Item ids are positive whole numbers; ` +
          "an id from createItem() may also be a provisional \"pending-<n>\" id.");
    }

    // Authorized for a provisional id too: whether a submission became an item, and what it says,
    // is a read of this list's state like any other.
    await this.#ctx.approvalQueue.authorizeObservation({
      title: sanitizeApprovalTitle(`Read item ${echoName(id)} from ${this.#ctx.listName()}`),
      description:
          "Read one item of this SharePoint list.\n\n" + formatApprovalField("Item id", id),
    });

    if (actionId !== null) return await this.#resolveProvisional(actionId);
    return await this.#ctx.api.getItem(this.#ctx.siteId, this.#ctx.listId, id);
  }

  /** What became of a submission, from its durable outcome record. */
  async #resolveProvisional(actionId: number): Promise<ListItem> {
    let outcome = this.#ctx.kv.get<ActionOutcome>(outcomeKey(actionId));
    if (outcome?.state === "applied") {
      return await this.#ctx.api.getItem(this.#ctx.siteId, this.#ctx.listId, outcome.itemId);
    }
    if (outcome) throw submissionError(actionId, outcome);

    // No record: either still queued, or stuck behind an earlier failure. The platform's drain stops
    // at a failed action rather than skipping it, so saying which one it is saves the submitter from
    // waiting on something that will never move on its own.
    if (blockedByOlderFailure(this.#ctx, actionId)) {
      throw new Error(
          `Submission ${actionId} is queued behind an earlier failed submission; a build ` +
          "collaborator must clear it from the workspace's activity list.");
    }
    throw new Error(`Submission ${actionId} has not been applied yet.`);
  }

  async createItem(fields: Record<string, unknown>): Promise<CreateItemResult> {
    let columns = await this.#ctx.columns();
    let validated = validateFields(fields, columns);

    if (this.#ctx.pendingActions.list().length >= MAX_PENDING_ACTIONS) {
      throw new Error(
          "Too many pending SharePoint list actions. Resolve existing actions before adding more.");
    }

    let actionId = this.#ctx.pendingActions.submit({ type: "createItem", fields: validated });

    // Registered before the action is submitted: with the auto-approval rule on, the apply can land
    // before submitAction() has even returned.
    let wait = this.#ctx.waitForOutcome(actionId);
    try {
      await this.#ctx.approvalQueue.submitAction(actionId, {
        title: sanitizeApprovalTitle(`Create item in ${this.#ctx.listName()}`),
        ...describeCreateItem(validated),
        actionKind: CREATE_ITEM_ACTION,
        // The author's verdict that adding a row to a list the user chose is safe to apply without
        // per-action review. The user still has to enable the rule for this list.
        autoApprovable: true,
        implementsRevert: false,
        // No `awaitDecision`: a pending create is reported honestly through a provisional id that
        // getItem() resolves, so the agent has something to keep working with.
      });
    } catch (err) {
      wait.cancel();
      this.#ctx.pendingActions.remove(actionId);
      throw err;
    }

    let outcome = await wait.outcome;
    if (outcome === null) return { id: provisionalItemId(actionId), provisional: true };
    if (outcome.state === "applied") return { id: outcome.itemId, provisional: false };
    throw submissionError(actionId, outcome);
  }
}

// =======================================================================================

export type SharePointListGatekeeperImplProps = {
  userObjectId: string;
  /** Graph site id, resolved from the pasted URL when the binding was introduced. */
  siteId: string;
  /** Graph list id. Immutable: a rename changes the list's URL and title, never this. */
  listId: string;
}

@validateRpc()
export class SharePointListGatekeeperImpl
    extends DurableObject<Env, SharePointListGatekeeperImplProps>
    implements Gatekeeper<SharePointListSession> {
  #tokens = new AccessTokenCache(opts => this.#account().getAccessToken(opts));
  #columnCache: ColumnDefinition[] | undefined;
  #outcomeWaiters = new Map<number, (outcome: ActionOutcome | null) => void>();
  #waitForApplyMs: number;

  constructor(
      ctx: DurableObjectState<SharePointListGatekeeperImplProps>, env: Env,
      waitForApplyMs: number = WAIT_FOR_APPLY_MS) {
    super(ctx, env);
    this.#waitForApplyMs = waitForApplyMs;
  }

  #account(): DurableObjectStub<UserAccount> {
    return this.ctx.exports.UserAccount.get(
        this.ctx.exports.UserAccount.idFromString(this.ctx.props.userObjectId));
  }

  #getAccessToken(opts?: AccessTokenRequest): Promise<string> {
    return this.#tokens.get(opts);
  }

  /**
   * A SharePoint client wired to this account.
   *
   * Same claims-challenge hook as the mailbox gatekeeper: the rejection is reported straight to the
   * account authority, and this object's own token memo is dropped first because it is the only copy
   * the authority cannot reach.
   *
   * Two ceilings, one factory. Reads fail fast on a throttle because an agent is waiting on them;
   * applies sit it out because an approval queue is waiting instead.
   */
  #api(retryAfter: RetryAfterPolicy): GraphSharePointApi {
    return new GraphSharePointApi(opts => this.#getAccessToken(opts), {
      retryAfter,
      onCredentialsRejected: async (detail: string) => {
        this.#tokens.invalidate();
        await this.#account().reportCredentialsRejected(detail);
      },
    });
  }

  /** The list's columns, cached for this object's lifetime. Invalidated by any 4xx from a create. */
  async #columns(): Promise<ColumnDefinition[]> {
    if (this.#columnCache) return this.#columnCache;
    let columns =
        await this.#api(READ_RETRY_AFTER).listColumns(this.ctx.props.siteId, this.ctx.props.listId);
    this.#columnCache = columns;
    return columns;
  }

  /** The list's name for display, from the cache `describe()` fills. */
  #listName(): string {
    return this.ctx.storage.kv.get<ListMeta>(META_KEY)?.displayName ?? "this SharePoint list";
  }

  /**
   * Wait for one action's outcome, for a bounded time.
   *
   * Best effort by design: the waiter lives in memory, so an eviction between submit and apply drops
   * it and the call answers with a provisional id. The durable outcome record is what actually
   * carries the answer; this only shortens the common case.
   */
  #waitForOutcome(actionId: number): OutcomeWait {
    if (this.#waitForApplyMs <= 0) {
      return { outcome: Promise.resolve(null), cancel: () => {} };
    }
    let settle: (outcome: ActionOutcome | null) => void;
    let outcome = new Promise<ActionOutcome | null>(resolve => { settle = resolve; });
    let timer = setTimeout(() => this.#settleWait(actionId, null), this.#waitForApplyMs);
    this.#outcomeWaiters.set(actionId, result => {
      clearTimeout(timer);
      settle(result);
    });
    return { outcome, cancel: () => this.#settleWait(actionId, null) };
  }

  #settleWait(actionId: number, outcome: ActionOutcome | null): void {
    let waiter = this.#outcomeWaiters.get(actionId);
    if (!waiter) return;
    this.#outcomeWaiters.delete(actionId);
    waiter(outcome);
  }

  async describe(): Promise<ResourceDescription> {
    let cached = this.ctx.storage.kv.get<ListMeta>(META_KEY);
    let meta: ListMeta;
    try {
      let info = await this.#api(READ_RETRY_AFTER)
          .describeList(this.ctx.props.siteId, this.ctx.props.listId);
      meta = { displayName: info.displayName, webUrl: info.webUrl };
      this.ctx.storage.kv.put(META_KEY, meta);
    } catch (err) {
      // The overseer tears the binding down when describe() throws, and it calls describe() again on
      // every listing — so a Graph blip after introduction must not cost the user their binding.
      // With nothing cached this is introduction time, where failing is the right answer.
      if (!cached) throw err;
      meta = cached;
    }

    return {
      url: meta.webUrl,
      title: meta.displayName,
      snippet: "SharePoint list",
      suggestedBindingName: "SHAREPOINT_LIST",
      tsType: "SharePointListSession",
    };
  }

  async getTypeScriptTypes(): Promise<string> {
    return SHAREPOINT_TYPES_CODE;
  }

  async getAutoApprovableActions(): Promise<ActionKind[]> {
    // Creating an item only adds a row to a list the user picked; it destroys nothing and changes
    // nothing that was already there. The rule is still off until the user turns it on, and it
    // applies to this list across the workspace.
    return [CREATE_ITEM_ACTION];
  }

  async startSession(approvalQueue: RpcStub<ApprovalQueue>): Promise<SharePointListSession> {
    return new SharePointListSessionImpl({
      api: this.#api(READ_RETRY_AFTER),
      approvalQueue: approvalQueue.dup(),
      pendingActions: new PendingActionStore<SharePointListAction>(this.ctx.storage.kv),
      kv: this.ctx.storage.kv,
      siteId: this.ctx.props.siteId,
      listId: this.ctx.props.listId,
      listName: () => this.#listName(),
      columns: () => this.#columns(),
      waitForOutcome: actionId => this.#waitForOutcome(actionId),
    });
  }

  // ---------------------------------------------------------------------------
  // The only place this gatekeeper writes to SharePoint.

  /**
   * Create the submitted item, idempotently.
   *
   * The overseer may call this again for an action it already applied — a crash between this method
   * returning and the overseer recording it looks exactly like an action that was never applied. The
   * outcome record is what tells the two apart, so it is written in the same transaction that clears
   * the pending action: either both landed or neither did.
   */
  async applyAction(actionId: number): Promise<void> {
    let kv = this.ctx.storage.kv;
    let pendingActions = new PendingActionStore<SharePointListAction>(kv);

    let recorded = kv.get<ActionOutcome>(outcomeKey(actionId));
    if (recorded?.state === "applied") {
      // The item exists. Finish the bookkeeping the crash interrupted rather than creating a second
      // copy of it.
      pendingActions.remove(actionId);
      this.#settleWait(actionId, recorded);
      return;
    }

    let action = pendingActions.get(actionId);
    if (!action) throw new Error(`Unknown pending SharePoint list action: ${actionId}`);

    let item: ListItem;
    try {
      item = await this.#api(APPLY_RETRY_AFTER)
          .createItem(this.ctx.props.siteId, this.ctx.props.listId, action.fields);
    } catch (err) {
      // A 4xx means the schema this create was validated against no longer matches the list.
      if (err instanceof GraphApiError && err.status >= 400 && err.status < 500) {
        this.#columnCache = undefined;
      }
      let outcome = failedOutcome(err);
      kv.put(outcomeKey(actionId), outcome);
      this.#settleWait(actionId, outcome);
      // Rethrown so the drain stops here and the action stays pending for a human to retry or
      // reject. Everything submitted after it waits behind it; getItem() says so.
      throw err;
    }

    let outcome: ActionOutcome = {
      state: "applied",
      itemId: item.id,
      ...(item.webUrl ? { webUrl: item.webUrl } : {}),
    };
    this.ctx.storage.transactionSync(() => {
      kv.put(outcomeKey(actionId), outcome);
      pendingActions.remove(actionId);
    });
    this.#settleWait(actionId, outcome);
  }

  async rejectAction(actionId: number): Promise<void | {restart?: boolean}> {
    let kv = this.ctx.storage.kv;
    let pendingActions = new PendingActionStore<SharePointListAction>(kv);
    if (!pendingActions.get(actionId)) {
      throw new Error(`Unknown pending SharePoint list action: ${actionId}`);
    }

    let outcome: ActionOutcome = { state: "rejected" };
    this.ctx.storage.transactionSync(() => {
      kv.put(outcomeKey(actionId), outcome);
      pendingActions.remove(actionId);
    });
    this.#settleWait(actionId, outcome);
  }

  async revertAction(_action: number):
      Promise<void | {message?: string, canRetry?: boolean, restart?: boolean}> {
    // Deleting the item would need Sites.Manage-level intent this connection never asked for, and
    // the row may already have been edited by someone else. Actions are submitted with
    // implementsRevert: false, so the UI never offers this.
    throw new Error("revert is not implemented");
  }

  /**
   * Observer admission: the collaborator must be able to open this list themselves.
   *
   * A list is one access unit — anything read through this gatekeeper came out of it — so a single
   * check answers for every past and future observation, and nothing needs to be tracked per
   * observer. The check runs on the observer's own Microsoft token, inside the verifier the overseer
   * minted for them, so this gatekeeper never sees their credentials and cannot be fooled by the
   * owner's access standing in for theirs.
   */
  async addObserver(_id: string, user: Fetcher<GatekeeperUserVerifier>): Promise<void> {
    let verifier = user as unknown as Fetcher<SharePointVerifierApi>;
    let allowed = await verifier.hasListAccess(this.ctx.props.siteId, this.ctx.props.listId);
    if (!allowed) throw new Error("You do not have access to this SharePoint list.");
  }

  async removeObserver(_id: string): Promise<void> {}
}
