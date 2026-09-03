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

/** Metadata for one attachment on a message. Never carries the attachment's content. */
export type OutlookAttachmentInfo = {
  /** Id of this attachment within its message. Pass it to `getAttachmentContent()`. */
  id: string;
  /** Filename as the sender supplied it. Untrusted text: it is chosen by whoever sent the mail. */
  name: string;
  /** MIME type Outlook recorded, e.g. "application/pdf". Also chosen by the sender. */
  mimeType: string;
  /**
   * Size in bytes as Outlook reports it. For a file attachment this measures the stored form,
   * which is a little larger than the file itself, so treat it as an upper bound.
   */
  sizeBytes: number;
  /** True for images and other parts displayed inside the message body rather than listed. */
  isInline: boolean;
  /**
   * What this attachment actually is:
   *
   * - `"file"` — an ordinary attached file. The only kind whose contents can be read.
   * - `"item"` — another Outlook item (email, event, contact) attached to this message.
   * - `"reference"` — a link to a file in OneDrive or SharePoint. The bytes live there, not in
   *   the mailbox.
   */
  kind: "file" | "item" | "reference";
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

  /**
   * Get a message capability for a message id seen earlier (see `OutlookMessageInfo.id`).
   *
   * Nothing is read and no id is checked here: any non-empty id is accepted, and a wrong one simply
   * fails on the first call made against it. Each read through the returned message is authorized
   * on its own, exactly as it is for a message obtained from a cursor.
   */
  getMessage(messageId: string): Promise<OutlookMessage>;
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
   * List this message's attachments. Metadata only — no content is downloaded, so this is cheap
   * even for a message carrying large files.
   *
   * One page: a message holding more attachments than Outlook returns at once is refused rather
   * than listed in part, so what comes back is always the message's complete set.
   */
  listAttachments(): Promise<OutlookAttachmentInfo[]>;

  /**
   * Read one attachment's contents, exactly as stored: nothing is converted, decoded or
   * interpreted here.
   *
   * Only `kind: "file"` attachments have contents; item and reference attachments throw, as does an
   * attachment over 10 MiB — which is refused from its metadata, before anything is downloaded.
   *
   * Treat the bytes as untrusted input: their name, type and content all come from whoever sent
   * the message.
   */
  getAttachmentContent(attachmentId: string): Promise<ArrayBuffer>;

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
