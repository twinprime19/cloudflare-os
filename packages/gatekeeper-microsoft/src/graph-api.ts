// Microsoft Graph mail client.
//
// Every mailbox request in this package goes through `GraphMailApi.#request()`. That single
// chokepoint is what guarantees the two properties the rest of the code depends on:
//
//   - `Prefer: IdType="ImmutableId"` on EVERY request. Outlook's default message ids change when a
//     message moves between folders, and this gatekeeper captures ids when an action is queued and
//     uses them again when the action is approved — possibly minutes later, after a user drag, an
//     inbox rule, or an earlier approved move has relocated the message. Immutable ids are the only
//     ids that still resolve then.
//   - Bearer injection and the 401 / claims-challenge / throttling policy, via `fetchWithAuthRetry`.
//
// URL construction is likewise centralized: ids and search terms reach this file from an agent, so
// path segments are percent-encoded and every query parameter goes through `URLSearchParams`.
// Interpolating either raw would let a crafted id escape into the path or into OData syntax.

import {
  AccessTokenProvider, CredentialsRejectedReporter, claimsChallengeDetail, fetchWithAuthRetry,
} from "./auth-retry";
import type { OutlookAttachmentInfo, OutlookFolderInfo, OutlookMessageInfo } from "./types";

/** Graph's origin. Pagination links are pinned to it; see `assertGraphUrl`. */
export const GRAPH_ORIGIN = "https://graph.microsoft.com";

const GRAPH_BASE = `${GRAPH_ORIGIN}/v1.0`;

/** Per-attempt ceiling on a Graph round trip. */
const GRAPH_TIMEOUT_MS = 20 * 1000;

/** Longest provider error text echoed into an error a user or agent can see. */
const MAX_GRAPH_ERROR_CHARS = 500;

/** Message fields every listing and metadata read selects. Keeps payloads bounded. */
const MESSAGE_SELECT = [
  "id", "subject", "from", "toRecipients", "ccRecipients", "receivedDateTime", "isRead",
  "hasAttachments", "bodyPreview", "parentFolderId", "webLink", "conversationId",
].join(",");

/**
 * Attachment fields every attachment read selects.
 *
 * Selecting is what keeps a listing cheap: without `$select`, Graph returns `contentBytes` for
 * every file attachment, so listing a message with a 9 MiB PDF would download the PDF.
 *
 * Only properties of the base `microsoft.graph.attachment` type may appear here — Graph rejects a
 * `$select` naming a property that exists solely on a derived type (`contentBytes`, `contentId`).
 * `@odata.type` is deliberately absent for the same reason: it is OData control information rather
 * than a property, and Graph emits it anyway, because the collection is declared as the base type
 * while every instance is one of the derived ones.
 */
const ATTACHMENT_SELECT = ["id", "name", "contentType", "size", "isInline"].join(",");

/**
 * Ceiling on a single attachment read, in bytes.
 *
 * Mirrors the workshop's 10 MiB document limit (`MAX_CONVERTIBLE_DOCUMENT_BYTES` in
 * workshop-backend). It is repeated here as a literal rather than imported: this gatekeeper is its
 * own worker and must bound its RPC payloads whatever the application on the other side does.
 */
export const MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024; // 10 MiB

// `childFolderCount` is what makes the folder walk cheap: a folder reporting none is never asked
// for its children.
const FOLDER_SELECT = [
  "id", "displayName", "parentFolderId", "totalItemCount", "unreadItemCount", "childFolderCount",
].join(",");

/** Page sizes Graph is asked for. `$top` above 100 is rejected for mail collections. */
export const DEFAULT_MESSAGE_PAGE_SIZE = 25;
const MAX_MESSAGE_PAGE_SIZE = 50;

/**
 * Ceilings on the folder walk, which pages internally rather than through a cursor. Three
 * independent bounds: pages within one collection, requests across the whole nested walk, and
 * folders returned.
 */
const FOLDER_PAGE_SIZE = 100;
const MAX_FOLDER_PAGES = 10;
const MAX_FOLDER_REQUESTS = 25;
const MAX_FOLDERS = 500;

/** Longest KQL search string accepted from an agent. */
const MAX_SEARCH_QUERY_CHARS = 400;

/** Longest reply body accepted, matching the cap the mail UIs impose in practice. */
export const MAX_REPLY_BODY_BYTES = 64 * 1024;

type GraphRecipient = {
  emailAddress?: { address?: string; name?: string };
};

type GraphMessage = {
  id?: string;
  subject?: string;
  from?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  bodyPreview?: string;
  parentFolderId?: string;
  webLink?: string;
  conversationId?: string;
  body?: { contentType?: string; content?: string };
};

type GraphMailFolder = {
  id?: string;
  displayName?: string;
  parentFolderId?: string;
  totalItemCount?: number;
  unreadItemCount?: number;
  childFolderCount?: number;
};

type GraphAttachment = {
  /** OData type annotation: `#microsoft.graph.fileAttachment` and friends. */
  "@odata.type"?: string;
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
};

type GraphCollection<T> = {
  value?: T[];
  "@odata.nextLink"?: string;
};

/** A page of results plus the link to the next one, when Graph offers it. */
export type GraphPage<T> = {
  items: T[];
  nextLink?: string;
};

/**
 * A Graph failure, classified by what the caller can do about it.
 *
 * `credentialsRejected` marks the claims-challenge case: the account has already been told its
 * credentials are dead, and no retry will help until the user reconnects.
 */
export class GraphApiError extends Error {
  readonly status: number;
  readonly code: string;
  readonly credentialsRejected: boolean;

  constructor(
      status: number, code: string, message: string, opts: { credentialsRejected?: boolean } = {}) {
    super(message);
    this.name = "GraphApiError";
    this.status = status;
    this.code = code;
    this.credentialsRejected = opts.credentialsRejected ?? false;
  }
}

function truncate(value: string, max = MAX_GRAPH_ERROR_CHARS): string {
  return value.length > max ? `${value.slice(0, max)}…` : value;
}

/**
 * Build a Graph URL from already-decoded path segments and query parameters.
 *
 * Segments are percent-encoded individually, so an id containing `/`, `?`, `#` or `..` stays one
 * segment instead of rewriting the request path. Every parameter value is percent-encoded, so an
 * OData operator or an `&` inside a search term cannot terminate the value it sits in or add a
 * parameter.
 *
 * Parameter NAMES are written literally, which is what keeps the OData system options readable as
 * `$select` rather than `%24select`. That is safe only because every name comes from a literal in
 * this file — no caller-supplied string ever becomes a parameter name.
 *
 * Empty and dot segments are rejected rather than encoded: percent-encoding leaves `.` and `..`
 * intact, and URL parsing collapses them at request time, so an id of `..` would retarget the
 * request one level up the path (`/me/mailFolders/..` becomes `/me`) with the bearer token
 * attached. Refusing them keeps "one segment stays one segment" literally true.
 */
export function graphUrl(segments: string[], params?: Record<string, string>): string {
  let path = segments.map(segment => {
    if (segment === "" || segment === "." || segment === "..") {
      throw new Error("Microsoft Graph path segments must not be empty or relative.");
    }
    return encodeURIComponent(segment);
  }).join("/");
  let query = Object.entries(params ?? {})
      .map(([name, value]) => `${name}=${encodeURIComponent(value)}`)
      .join("&");
  return `${GRAPH_BASE}/${path}${query ? `?${query}` : ""}`;
}

/**
 * Return `url` if it is a Graph URL, otherwise throw.
 *
 * Applied to every `@odata.nextLink` before it is followed. The fetch helper attaches this
 * mailbox's bearer token to whatever URL it is given, so following a link to another origin would
 * hand the token to that host. Graph controls the link today; this makes a compromised or spoofed
 * response unable to exfiltrate the token regardless.
 */
export function assertGraphUrl(url: string): string {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    throw new Error("Microsoft Graph returned a malformed pagination link.");
  }
  if (parsed.origin !== GRAPH_ORIGIN) {
    throw new Error(
        `Microsoft Graph returned a pagination link outside ${GRAPH_ORIGIN}; refusing to follow it.`);
  }
  return parsed.toString();
}

function addressFrom(recipient: GraphRecipient | undefined) {
  let address = recipient?.emailAddress?.address;
  if (typeof address !== "string" || !address) return undefined;
  let name = recipient?.emailAddress?.name;
  return {
    address,
    ...(typeof name === "string" && name ? { name } : {}),
  };
}

function addressesFrom(recipients: GraphRecipient[] | undefined) {
  if (!Array.isArray(recipients)) return [];
  return recipients
      .map(addressFrom)
      .filter((entry): entry is { address: string; name?: string } => entry !== undefined);
}

function messageInfoFrom(message: GraphMessage): OutlookMessageInfo {
  if (typeof message.id !== "string" || !message.id) {
    throw new Error("Microsoft Graph returned a message without an id.");
  }
  let received = message.receivedDateTime ? new Date(message.receivedDateTime) : new Date(0);
  let from = addressFrom(message.from);
  return {
    id: message.id,
    subject: message.subject ?? "(no subject)",
    ...(from ? { from } : {}),
    to: addressesFrom(message.toRecipients),
    cc: addressesFrom(message.ccRecipients),
    receivedAt: Number.isNaN(received.valueOf()) ? new Date(0) : received,
    isRead: message.isRead === true,
    hasAttachments: message.hasAttachments === true,
    preview: message.bodyPreview ?? "",
    ...(message.parentFolderId ? { folderId: message.parentFolderId } : {}),
    ...(message.webLink ? { webLink: message.webLink } : {}),
    ...(message.conversationId ? { conversationId: message.conversationId } : {}),
  };
}

function folderInfoFrom(folder: GraphMailFolder): OutlookFolderInfo {
  if (typeof folder.id !== "string" || !folder.id) {
    throw new Error("Microsoft Graph returned a mail folder without an id.");
  }
  return {
    id: folder.id,
    name: folder.displayName ?? "(unnamed folder)",
    ...(folder.parentFolderId ? { parentFolderId: folder.parentFolderId } : {}),
    totalItemCount: typeof folder.totalItemCount === "number" ? folder.totalItemCount : 0,
    unreadItemCount: typeof folder.unreadItemCount === "number" ? folder.unreadItemCount : 0,
  };
}

/**
 * Classify an attachment from its OData type annotation.
 *
 * Anything not recognized as a file or a reference is reported as `"item"`, the kind that carries
 * no downloadable bytes. That is the safe default: a type this code has never seen is not something
 * `/$value` can be assumed to serve, and misreporting it as a file would turn a Graph error into an
 * unexplained download failure.
 */
function attachmentKind(odataType: string | undefined): OutlookAttachmentInfo["kind"] {
  let type = (odataType ?? "").replace(/^#/, "").toLowerCase();
  if (type === "microsoft.graph.fileattachment") return "file";
  if (type === "microsoft.graph.referenceattachment") return "reference";
  return "item";
}

function attachmentInfoFrom(attachment: GraphAttachment): OutlookAttachmentInfo {
  if (typeof attachment.id !== "string" || !attachment.id) {
    throw new Error("Microsoft Graph returned an attachment without an id.");
  }
  return {
    id: attachment.id,
    name: attachment.name ?? "(unnamed attachment)",
    mimeType: attachment.contentType ?? "application/octet-stream",
    sizeBytes: typeof attachment.size === "number" && attachment.size > 0 ? attachment.size : 0,
    isInline: attachment.isInline === true,
    kind: attachmentKind(attachment["@odata.type"]),
  };
}

/** An attachment name as it can safely appear in a one-line error message. */
function attachmentLabel(name: string): string {
  return truncate(name.replace(/[\r\n]+/g, " "), 100);
}

/**
 * Validate an agent-supplied KQL search string.
 *
 * A double quote would close the quoted KQL literal the term is wrapped in, turning the rest of the
 * term into search syntax; rejecting it keeps the value a value. The length cap keeps a pathological
 * query from becoming a Graph error loop.
 */
export function validateSearchQuery(query: string): string {
  let trimmed = query.trim();
  if (!trimmed) throw new Error("Search query must not be empty.");
  if (trimmed.length > MAX_SEARCH_QUERY_CHARS) {
    throw new Error(`Search query must be at most ${MAX_SEARCH_QUERY_CHARS} characters.`);
  }
  if (trimmed.includes("\"")) {
    throw new Error("Search query must not contain double quotes.");
  }
  return trimmed;
}

function pageSize(requested: number | undefined): string {
  let size = requested ?? DEFAULT_MESSAGE_PAGE_SIZE;
  if (!Number.isFinite(size)) size = DEFAULT_MESSAGE_PAGE_SIZE;
  return String(Math.min(Math.max(Math.trunc(size), 1), MAX_MESSAGE_PAGE_SIZE));
}

export type GraphMailApiOptions = {
  /** Called when Graph answers a claims challenge, so the account can be marked dead. */
  onCredentialsRejected?: CredentialsRejectedReporter;
};

export class GraphMailApi {
  #getAccessToken: AccessTokenProvider;
  #onCredentialsRejected: CredentialsRejectedReporter | undefined;

  constructor(getAccessToken: AccessTokenProvider, opts: GraphMailApiOptions = {}) {
    this.#getAccessToken = getAccessToken;
    this.#onCredentialsRejected = opts.onCredentialsRejected;
  }

  /**
   * The one place a Graph request is made.
   *
   * `Prefer: IdType="ImmutableId"` is set here, unconditionally, so no call site can forget it. The
   * text body preference is added to reads, where it is what Graph honors: it makes message bodies
   * come back as plain text rather than HTML that would have to be sanitized before an agent sees
   * it.
   */
  async #request(url: string, init: RequestInit = {}): Promise<Response> {
    let method = (init.method ?? "GET").toUpperCase();
    let headers = new Headers(init.headers);
    let prefer = ['IdType="ImmutableId"'];
    if (method === "GET") prefer.push('outlook.body-content-type="text"');
    headers.set("Prefer", prefer.join(", "));
    if (!headers.has("Accept")) headers.set("Accept", "application/json");
    if (init.body && !headers.has("Content-Type")) headers.set("Content-Type", "application/json");

    return await fetchWithAuthRetry(url, { ...init, headers }, this.#getAccessToken, {
      timeoutMs: GRAPH_TIMEOUT_MS,
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

  /** Read a response body as bytes. Used only for `/$value`, which serves no JSON. */
  async #fetchBytes(url: string): Promise<ArrayBuffer> {
    let response = await this.#request(url, { headers: { Accept: "*/*" } });
    if (!response.ok) throw await this.#toError(response);
    return await response.arrayBuffer();
  }

  async #send(url: string, init: RequestInit): Promise<void> {
    let response = await this.#request(url, init);
    if (!response.ok) throw await this.#toError(response);
    await response.body?.cancel();
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
            "Microsoft requires this account to sign in again before the mailbox can be used " +
            `(${claims}). Please reconnect the account.`,
            { credentialsRejected: true });
      }
      return new GraphApiError(401, code || "unauthenticated",
          `Microsoft rejected the mailbox credentials (${detail}).`);
    }

    if (response.status === 403) {
      return new GraphApiError(403, code || "forbidden",
          "This Microsoft connection is not permitted to perform that mailbox operation " +
          `(${detail}). Reconnect the account and grant mailbox access.`);
    }

    if (response.status === 404) {
      return new GraphApiError(404, code || "notFound",
          `That mailbox item no longer exists (${detail}).`);
    }

    if (response.status === 429) {
      let retryAfter = response.headers.get("Retry-After");
      return new GraphApiError(429, code || "activityLimitReached",
          "Microsoft Graph is throttling this mailbox" +
          (retryAfter ? `; retry in ${truncate(retryAfter, 20)} seconds` : "") + ".");
    }

    if (response.status >= 500) {
      return new GraphApiError(response.status, code || "serviceError",
          `Microsoft Graph is temporarily unavailable (${detail}).`);
    }

    return new GraphApiError(response.status, code || "unknown",
        `Microsoft Graph request failed: ${response.status}${code ? ` ${code}` : ""} — ${detail}`);
  }

  // ── Reads ─────────────────────────────────────────────────────────────────────

  /**
   * Every mail folder in the mailbox, including nested ones.
   *
   * `GET /me/mailFolders` returns only the root folder's children, so each folder that reports
   * children is visited in turn — otherwise a message could never be moved into a subfolder,
   * because its id would never be discoverable. The walk is breadth-first (shallow folders, the
   * ones a user actually names, are found first) and bounded three ways: total requests, pages per
   * collection, and folders returned. Folders reporting no children are never visited, so a flat
   * mailbox costs exactly the requests the old root-only listing did.
   */
  async listFolders(): Promise<OutlookFolderInfo[]> {
    let folders: OutlookFolderInfo[] = [];
    // `undefined` stands for the mailbox root, whose children come from a different path.
    let pending: (string | undefined)[] = [undefined];
    let requests = 0;

    while (pending.length > 0 && requests < MAX_FOLDER_REQUESTS && folders.length < MAX_FOLDERS) {
      let parentId = pending.shift();
      let url: string | undefined = parentId === undefined
          ? graphUrl(["me", "mailFolders"], {
              $select: FOLDER_SELECT,
              $top: String(FOLDER_PAGE_SIZE),
              // Hidden folders (recoverable items, sync artifacts) are not places to file mail.
              includeHiddenFolders: "false",
            })
          : graphUrl(["me", "mailFolders", parentId, "childFolders"], {
              $select: FOLDER_SELECT,
              $top: String(FOLDER_PAGE_SIZE),
            });

      for (let page = 0;
           url && page < MAX_FOLDER_PAGES && requests < MAX_FOLDER_REQUESTS;
           page++) {
        requests++;
        let body: GraphCollection<GraphMailFolder> =
            await this.#fetchJson<GraphCollection<GraphMailFolder>>(url);
        for (let folder of body.value ?? []) {
          if (folders.length >= MAX_FOLDERS) return folders;
          folders.push(folderInfoFrom(folder));
          if (folder.id && typeof folder.childFolderCount === "number" &&
              folder.childFolderCount > 0) {
            pending.push(folder.id);
          }
        }
        let next = body["@odata.nextLink"];
        url = next ? assertGraphUrl(next) : undefined;
      }
    }
    return folders;
  }

  async getFolder(folderId: string): Promise<OutlookFolderInfo> {
    let folder = await this.#fetchJson<GraphMailFolder>(
        graphUrl(["me", "mailFolders", folderId], { $select: FOLDER_SELECT }));
    return folderInfoFrom(folder);
  }

  /**
   * First page of messages, newest first. Omit `folderId` to read across the whole mailbox.
   */
  async listMessages(opts: { folderId?: string; pageSize?: number } = {})
      : Promise<GraphPage<OutlookMessageInfo>> {
    let segments = opts.folderId
        ? ["me", "mailFolders", opts.folderId, "messages"]
        : ["me", "messages"];
    return await this.#messagePage(graphUrl(segments, {
      $select: MESSAGE_SELECT,
      $top: pageSize(opts.pageSize),
      $orderby: "receivedDateTime desc",
    }));
  }

  /**
   * First page of messages matching a KQL search.
   *
   * `$orderby` is deliberately absent: Graph rejects it alongside `$search`, and results come back
   * in relevance order instead.
   */
  async searchMessages(query: string, opts: { pageSize?: number } = {})
      : Promise<GraphPage<OutlookMessageInfo>> {
    let validated = validateSearchQuery(query);
    return await this.#messagePage(graphUrl(["me", "messages"], {
      $select: MESSAGE_SELECT,
      $top: pageSize(opts.pageSize),
      $search: `"${validated}"`,
    }));
  }

  /** Follow a `@odata.nextLink` produced by a previous page. */
  async nextMessagePage(nextLink: string): Promise<GraphPage<OutlookMessageInfo>> {
    return await this.#messagePage(assertGraphUrl(nextLink));
  }

  async #messagePage(url: string): Promise<GraphPage<OutlookMessageInfo>> {
    let body = await this.#fetchJson<GraphCollection<GraphMessage>>(url);
    let next = body["@odata.nextLink"];
    return {
      items: (body.value ?? []).map(messageInfoFrom),
      // Validated on arrival, not only when followed, so a bad link fails the page that produced it.
      ...(next ? { nextLink: assertGraphUrl(next) } : {}),
    };
  }

  async getMessage(messageId: string): Promise<OutlookMessageInfo> {
    let message = await this.#fetchJson<GraphMessage>(
        graphUrl(["me", "messages", messageId], { $select: MESSAGE_SELECT }));
    return messageInfoFrom(message);
  }

  /** The message body as plain text (see the text body preference on reads). */
  async getMessageBody(messageId: string): Promise<string> {
    let message = await this.#fetchJson<GraphMessage>(
        graphUrl(["me", "messages", messageId], { $select: "id,body" }));
    return typeof message.body?.content === "string" ? message.body.content : "";
  }

  /**
   * Metadata for every attachment on a message, without any content.
   *
   * One request, no pagination loop: a message holds a handful of attachments and Graph returns
   * them in a single page. A `@odata.nextLink` here would mean a message with more attachments than
   * Outlook itself permits, and paging a mailbox on the caller's behalf is not this method's job --
   * so such a listing is refused rather than truncated. A short answer must never be mistakable for
   * a complete one, by the caller or by the audit trail that records what was listed.
   */
  async listAttachments(messageId: string): Promise<OutlookAttachmentInfo[]> {
    let body = await this.#fetchJson<GraphCollection<GraphAttachment>>(
        graphUrl(["me", "messages", messageId, "attachments"], { $select: ATTACHMENT_SELECT }));
    if (body["@odata.nextLink"]) {
      throw new Error(
          "This message carries more attachments than Outlook returns in one page, so they " +
          "cannot be listed completely here. Open the message in Outlook to see them all.");
    }
    return (body.value ?? []).map(attachmentInfoFrom);
  }

  /** Metadata for one attachment, without its content. */
  async #getAttachment(messageId: string, attachmentId: string): Promise<OutlookAttachmentInfo> {
    let attachment = await this.#fetchJson<GraphAttachment>(
        graphUrl(["me", "messages", messageId, "attachments", attachmentId],
                 { $select: ATTACHMENT_SELECT }));
    return attachmentInfoFrom(attachment);
  }

  /**
   * One file attachment's raw bytes, with the metadata that describes them.
   *
   * Metadata is read first so a kind that has no bytes, or a file over the cap, costs one small
   * request instead of a download that is then thrown away. Graph's `size` counts the stored
   * (base64) form of a file attachment, so it slightly overstates the file — a file just under the
   * cap can therefore be refused here. No compensation math: the post-download check below is the
   * authoritative one, and refusing marginally early is the safe direction.
   *
   * The body is buffered rather than streamed: the size gate above has already bounded it, and
   * Graph is a first-party origin whose declared size the callers are entitled to trust.
   */
  async getAttachmentBytes(messageId: string, attachmentId: string)
      : Promise<{ info: OutlookAttachmentInfo; content: ArrayBuffer }> {
    let info = await this.#getAttachment(messageId, attachmentId);
    let label = attachmentLabel(info.name);

    if (info.kind === "item") {
      throw new Error(
          `"${label}" is an Outlook item attached to this message (an email, event or contact), ` +
          "not a file, so it has no file contents to read.");
    }
    if (info.kind === "reference") {
      throw new Error(
          `"${label}" is a link to a file stored in OneDrive or SharePoint, not a copy of the ` +
          "file, so the mailbox holds no contents to read. Open the link in Outlook instead.");
    }
    if (info.sizeBytes > MAX_ATTACHMENT_BYTES) {
      throw new Error(
          `"${label}" is ${info.sizeBytes} bytes, which is over the ${MAX_ATTACHMENT_BYTES}-byte ` +
          "limit for reading an attachment. It was not downloaded.");
    }

    // `$value` is written literally for the same reason the OData query options are: it is a
    // fixed segment from this file, not caller data, and percent-encoding it would ask Graph for a
    // child named "%24value". The id before it is still encoded by graphUrl().
    let content = await this.#fetchBytes(
        `${graphUrl(["me", "messages", messageId, "attachments", attachmentId])}/$value`);

    if (content.byteLength > MAX_ATTACHMENT_BYTES) {
      throw new Error(
          `"${label}" returned ${content.byteLength} bytes, which is over the ` +
          `${MAX_ATTACHMENT_BYTES}-byte limit for reading an attachment.`);
    }
    return { info, content };
  }

  // ── Writes ────────────────────────────────────────────────────────────────────
  //
  // Only the gatekeeper's approved-action executor calls these. Session methods queue actions and
  // return nothing; nothing below runs before a human (or policy) approves.

  async setMessageRead(messageId: string, read: boolean): Promise<void> {
    await this.#send(graphUrl(["me", "messages", messageId]), {
      method: "PATCH",
      body: JSON.stringify({ isRead: read }),
    });
  }

  /**
   * Move a message to another folder. The message keeps its immutable id, which is why a queued
   * action still resolves after an unrelated move.
   */
  async moveMessage(messageId: string, destinationFolderId: string): Promise<void> {
    await this.#send(graphUrl(["me", "messages", messageId, "move"]), {
      method: "POST",
      body: JSON.stringify({ destinationId: destinationFolderId }),
    });
  }

  /**
   * Create a reply draft in the mailbox's Drafts folder, with `comment` inserted above the quoted
   * original. Nothing is sent: the user reviews and sends the draft from Outlook.
   */
  async createReplyDraft(messageId: string, comment: string, replyAll: boolean): Promise<void> {
    if (new TextEncoder().encode(comment).byteLength > MAX_REPLY_BODY_BYTES) {
      throw new Error(`Reply body must be at most ${MAX_REPLY_BODY_BYTES} bytes.`);
    }
    let action = replyAll ? "createReplyAll" : "createReply";
    await this.#send(graphUrl(["me", "messages", messageId, action]), {
      method: "POST",
      body: JSON.stringify({ comment }),
    });
  }
}
