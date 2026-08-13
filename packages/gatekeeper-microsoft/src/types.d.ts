import { Cursor } from "@gadgets/workshop-shared/gatekeeper";

export type { Cursor };

// ── Plain data types ────────────────────────────────────────────────

/** An email address with optional display name. */
export type OutlookAddress = {
  address: string;
  name?: string;
}

/** Metadata for a single mailbox message. */
export type OutlookMessageInfo = {
  /**
   * Stable message id. It is an immutable id: it keeps identifying the same message even after the
   * message moves to another folder, so it stays valid across queued actions.
   */
  id: string;
  subject: string;
  /** Absent for messages with no sender, e.g. some drafts. */
  from?: OutlookAddress;
  to: OutlookAddress[];
  cc: OutlookAddress[];
  receivedAt: Date;
  isRead: boolean;
  hasAttachments: boolean;
  /** Short plain-text preview of the body. Use `getBody()` for the full text. */
  preview: string;
  /** Id of the folder currently holding the message, when Outlook reports one. */
  folderId?: string;
  /** Link that opens this message in Outlook on the web. */
  webLink?: string;
  /** Id shared by every message in the same conversation. */
  conversationId?: string;
}

/** Metadata for a mail folder. */
export type OutlookFolderInfo = {
  id: string;
  name: string;
  /** Absent for top-level folders. */
  parentFolderId?: string;
  totalItemCount: number;
  unreadItemCount: number;
}

// ── Capability interfaces ───────────────────────────────────────────
// These are RPC stubs — all methods are async. Capabilities can be
// passed across Worker boundaries and retain their access rights.

/** A single entry from a message cursor: metadata plus a message capability. */
export type OutlookMessageEntry = {
  info: OutlookMessageInfo;
  message: OutlookMessage;
}

/** A single entry from `listFolders()`: metadata plus a folder capability. */
export type OutlookFolderEntry = {
  info: OutlookFolderInfo;
  folder: OutlookFolder;
}

export interface OutlookMailSession {
  /**
   * List the most recent messages in the mailbox, newest first. Returns a cursor that lazily
   * fetches pages as they are consumed.
   */
  listMessages(): Promise<Cursor<OutlookMessageEntry>>;

  /**
   * Search the mailbox.
   *
   * `query` is Microsoft Search (KQL), not Gmail syntax: bare words match anywhere in the message,
   * and `property:value` terms are supported — e.g. `from:bob@example.com`, `subject:invoice`,
   * `hasAttachments:true`, `received>=2026-01-01`. Terms may be combined with `AND`/`OR`/`NOT`.
   * Double quotes are not accepted, since the query is itself sent as a quoted KQL string.
   *
   * Results come back in relevance order, NOT newest-first: Microsoft Graph refuses to sort a
   * search. Use `listMessages()` when recency matters.
   */
  search(query: string): Promise<Cursor<OutlookMessageEntry>>;

  /**
   * List the mailbox's folders, so messages can be moved to one of them.
   *
   * Nested folders are included: the result is the folder tree flattened, shallowest first, with
   * `parentFolderId` linking a folder to its parent. Very large or deeply nested mailboxes may be
   * truncated.
   */
  listFolders(): Promise<OutlookFolderEntry[]>;
}

export interface OutlookFolder {
  /** Get this folder's metadata (name, message counts). */
  getInfo(): Promise<OutlookFolderInfo>;

  /**
   * List the most recent messages in this folder, newest first. Returns a cursor that lazily
   * fetches pages as they are consumed.
   */
  listMessages(): Promise<Cursor<OutlookMessageEntry>>;
}

export interface OutlookMessage {
  /** Get message metadata (sender, recipients, subject, timestamps, read state). */
  getMetadata(): Promise<OutlookMessageInfo>;

  /** Get the full message body as plain text. */
  getBody(): Promise<string>;

  /**
   * Mark this message as read.
   *
   * Queued for approval: nothing changes in the mailbox until a human (or an auto-approval policy)
   * approves it, and the result is not observable from this session — later reads still report the
   * mailbox as it is now. Returns as soon as the request is queued.
   */
  markRead(): Promise<void>;

  /**
   * Mark this message as unread.
   *
   * Queued for approval, with the same semantics as `markRead()`.
   */
  markUnread(): Promise<void>;

  /**
   * Move this message to the folder with the given id (see `listFolders()`).
   *
   * Queued for approval: the message stays where it is until the action is approved, and this call
   * reports nothing about the outcome. The message keeps the same id after a move.
   */
  moveToFolder(folderId: string): Promise<void>;

  /**
   * Create a draft reply to the sender, with `body` above the quoted original.
   *
   * Queued for approval, and nothing is ever sent: once approved, the draft appears in the
   * mailbox's Drafts folder for the user to review and send. The draft is not readable from this
   * session and no draft id is returned.
   */
  createReplyDraft(body: string): Promise<void>;

  /**
   * Create a draft reply to the sender and every other recipient.
   *
   * Queued for approval, with the same semantics as `createReplyDraft()`.
   */
  createReplyAllDraft(body: string): Promise<void>;
}
