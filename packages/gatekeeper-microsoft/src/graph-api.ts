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
import type { OutlookFolderInfo, OutlookMessageInfo } from "./types";

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

const FOLDER_SELECT = [
  "id", "displayName", "parentFolderId", "totalItemCount", "unreadItemCount",
].join(",");

/** Page sizes Graph is asked for. `$top` above 100 is rejected for mail collections. */
export const DEFAULT_MESSAGE_PAGE_SIZE = 25;
const MAX_MESSAGE_PAGE_SIZE = 50;

/** Ceilings on the folder walk, which pages internally rather than through a cursor. */
const FOLDER_PAGE_SIZE = 100;
const MAX_FOLDER_PAGES = 10;
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
 */
export function graphUrl(segments: string[], params?: Record<string, string>): string {
  let path = segments.map(segment => encodeURIComponent(segment)).join("/");
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

  /** Every mail folder in the mailbox, paged internally under a hard cap. */
  async listFolders(): Promise<OutlookFolderInfo[]> {
    let folders: OutlookFolderInfo[] = [];
    let url: string | undefined = graphUrl(["me", "mailFolders"], {
      $select: FOLDER_SELECT,
      $top: String(FOLDER_PAGE_SIZE),
      // Mailboxes nest folders arbitrarily; the flat listing is what an agent can act on.
      includeHiddenFolders: "false",
    });

    for (let page = 0; url && page < MAX_FOLDER_PAGES; page++) {
      let body: GraphCollection<GraphMailFolder> =
          await this.#fetchJson<GraphCollection<GraphMailFolder>>(url);
      for (let folder of body.value ?? []) {
        if (folders.length >= MAX_FOLDERS) return folders;
        folders.push(folderInfoFrom(folder));
      }
      let next = body["@odata.nextLink"];
      url = next ? assertGraphUrl(next) : undefined;
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
