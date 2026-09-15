// Agent-facing types for the SharePoint list resource.
//
// `Cursor` is deliberately neither imported nor re-exported here. The vendor serves its whole type
// surface by concatenating this file after the Outlook mail types, and a second `Cursor`
// import/re-export would leave the combined text with a duplicate identifier. `SharePointCursor`
// reaches the same type through an inline import, so this file still stands on its own when the
// SharePoint list resource serves only its own types.
//
// Column shapes mirror what the Graph client normalises a list's schema into; nothing here is a
// SharePoint wire shape.

/** A pagination cursor. Call `next()` until it returns `null`; an empty batch is not exhaustion. */
export type SharePointCursor<T> = import("@gadgets/workshop-shared/gatekeeper").Cursor<T>;

// ── Plain data types ────────────────────────────────────────────────

/**
 * What a column holds.
 *
 * `person` and `lookup` are described so a form can show them, but they can never be written: both
 * are stored as a numeric directory/list id that a form has no way to supply. `unsupported` covers
 * every other column kind SharePoint offers (currency, calculated, hyperlink, managed metadata,
 * geolocation, …): readable as whatever SharePoint returns, never writable.
 */
export type ColumnType =
  | "text" | "number" | "boolean" | "dateTime" | "choice" | "person" | "lookup" | "unsupported";

/** One column of the list. */
export type ColumnDefinition = {
  /** Internal name. The one to use as a key in `createItem()`, in `select`, and in `where`. */
  name: string;
  /** Display name. Renameable by a list owner, so never use it to address a column. */
  displayName: string;
  type: ColumnType;
  /** `createItem()` refuses a call that omits this column. */
  required: boolean;
  /** SharePoint computes this column; it can be read but never written. */
  readOnly: boolean;
  /** Allowed values, for `choice` columns. `createItem()` refuses anything else. */
  choices?: string[];
  /** A `text` column that accepts more than one line. */
  multiline?: boolean;
  /** A `person` or `lookup` column that holds more than one value. */
  multiple?: boolean;
}

/** One item of the list. */
export type ListItem = {
  /** SharePoint's item id: a positive whole number as a string. Pass it to `getItem()`. */
  id: string;
  /** Link that opens this item in SharePoint. */
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  /** Column values, keyed by internal column name. */
  fields: Record<string, unknown>;
}

/** Comparisons `getItems()` accepts. `startswith` is text-only. */
export type WhereOperator = "eq" | "ne" | "lt" | "gt" | "le" | "ge" | "startswith";

/**
 * One filter clause. There is no way to pass a raw query string: the column is matched against the
 * list's real columns and the value is typed and escaped for you.
 */
export type WhereClause = {
  /** Internal column name, as `getColumns()` reports it. */
  column: string;
  op: WhereOperator;
  value: string | number | boolean;
}

/** How `getItems()` narrows what it walks. Every part is optional. */
export type GetItemsOptions = {
  /** Internal column names to read back. Defaults to every column. */
  select?: string[];
  /** Clauses combined with `and`. At most five; more than that wants a SharePoint view. */
  where?: WhereClause[];
  /** Items per page. Clamped to the list client's ceiling. */
  top?: number;
}

/** What `createItem()` returns. */
export type CreateItemResult = {
  /**
   * The new item's SharePoint id when the item was already created by the time the call returned,
   * otherwise a provisional id of the form `pending-<n>`.
   */
  id: string;
  /**
   * True when `id` is provisional: the item has been submitted but not yet written to SharePoint.
   * Pass the provisional id to `getItem()` later to find out what happened to it — it resolves to
   * the real item once applied, and throws a `SharePointSubmissionError` if the submission failed
   * or was rejected.
   */
  provisional: boolean;
}

/** Why a submission is not going to become an item. */
export type SharePointSubmissionOutcome = "failed" | "rejected";

/**
 * Thrown when a submitted item was decided against rather than written: `createItem()` throws it
 * when that happens while the call is still waiting, and `getItem()` throws it for a provisional id
 * whose submission ended the same way.
 */
export interface SharePointSubmissionError extends Error {
  name: "SharePointSubmissionError";
  outcome: SharePointSubmissionOutcome;
  /** True when SharePoint throttled the write, so an approver can simply retry it. */
  throttled: boolean;
}

// ── Capability interfaces ───────────────────────────────────────────
// These are RPC stubs — all methods are async. Capabilities can be
// passed across Worker boundaries and retain their access rights.

/**
 * One SharePoint list.
 *
 * Reads are audited; creating an item is submitted for approval. If the workspace has enabled the
 * "Create list items" auto-approval rule for this list, an uncontended `createItem()` usually
 * returns the real item id; otherwise it returns a provisional id and `getItem()` reports what
 * happened once the submission is decided.
 */
export interface SharePointListSession {
  /** The list's columns: their internal names, types, choices and required flags. */
  getColumns(): Promise<ColumnDefinition[]>;

  /** A cursor over the list's items, optionally filtered and narrowed. */
  getItems(options?: GetItemsOptions): Promise<SharePointCursor<ListItem>>;

  /**
   * One item by id.
   *
   * Accepts either a SharePoint item id or a provisional id from `createItem()`. A provisional id
   * is only meaningful to the session that issued it.
   */
  getItem(id: string): Promise<ListItem>;

  /**
   * Submit a new item.
   *
   * `fields` is keyed by internal column name. Every value is validated against the list's schema
   * before anything is submitted: unknown columns, missing required columns, wrong value types,
   * choices outside the column's options, and person/lookup columns are all refused here.
   */
  createItem(fields: Record<string, unknown>): Promise<CreateItemResult>;
}
