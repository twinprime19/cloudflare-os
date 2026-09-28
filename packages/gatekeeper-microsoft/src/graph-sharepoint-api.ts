// Microsoft Graph SharePoint list client.
//
// Same shape as the mailbox client next door (`graph-api.ts`): one request chokepoint, URLs built
// only by `graphUrl` (plus `graphSiteByPathUrl`, which is the one site address `graphUrl` cannot
// express), pagination links pinned to Graph's origin, bounded pages, truncated provider text. What
// differs is SharePoint-specific and deliberate:
//
//   - No `Prefer` header. The immutable-id and text-body preferences are Outlook options, and the
//     non-indexed-query override (`HonorNonIndexedQueriesWarningMayFailRandomly`) is deliberately
//     not sent: Microsoft's own wording for it is "may fail randomly", so a filter on an unindexed
//     column in a large list is left to fail with Graph's explanation, which names the column and
//     tells the owner to index it.
//   - A decorated `User-Agent`. SharePoint Online asks every app to identify itself and
//     deprioritizes undecorated traffic under load.
//   - A `Retry-After` ceiling supplied by the caller. SharePoint throttles against a per-minute
//     resource-unit budget and routinely asks for a minute or more, while the shared backoff clamps
//     every wait to ten seconds and replays straight back into the throttle. Reads therefore fail
//     fast beyond a short ceiling; the apply path, which has an approval queue behind it, can afford
//     a longer one, so the policy is a constructor option rather than a constant.
//
// No caller string ever reaches a URL path or an OData filter raw. Ids and names are path segments,
// which `graphUrl` encodes; filters are built from a structured `WhereClause[]` validated against
// the list's real columns, so there is no way to hand this client an OData expression.

import {
  AccessTokenProvider, CredentialsRejectedReporter, RetryAfterPolicy, claimsChallengeDetail,
  fetchWithAuthRetry,
} from "./auth-retry";
import {
  GRAPH_TIMEOUT_MS, GraphApiError, GraphCollection, GraphPage, assertGraphUrl, graphSiteByPathUrl,
  graphUrl, truncate,
} from "./graph-api";

/**
 * How this worker identifies itself to SharePoint Online, in the `NONISV|Company|App/Version` form
 * Microsoft documents for internal line-of-business apps. Undecorated traffic is deprioritized when
 * the service is under load, so this is set on every request rather than per call site.
 */
export const SHAREPOINT_USER_AGENT = "NONISV|Contoso|CloudflareOS/1.0";

/** Items Graph is asked for per page, and the ceiling a caller can raise it to. */
export const DEFAULT_ITEM_PAGE_SIZE = 25;
const MAX_ITEM_PAGE_SIZE = 200;

/** Pages of `/lists` walked while looking for the list a pasted URL names. */
const MAX_LIST_PAGES = 10;
const LIST_PAGE_SIZE = 100;

/** Clauses a single `$filter` may carry. Past this, the caller wants a view, not a query. */
const MAX_WHERE_CLAUSES = 5;

/** Longest text value accepted in a filter clause. Keeps a pathological value out of the URL. */
const MAX_FILTER_TEXT_CHARS = 256;

/** Longest caller-supplied name echoed back in a validation error. */
const MAX_NAME_ECHO_CHARS = 100;

/**
 * Default ceiling on a throttling wait. Matches the read path, which is what an agent is waiting on
 * synchronously; a caller with an approval queue behind it passes a longer one.
 */
const DEFAULT_MAX_RETRY_AFTER_MS = 10 * 1000;

/**
 * Columns SharePoint puts on every list that are plumbing rather than data. Hidden and read-only
 * columns are dropped by their own flags; these are the ones that are neither, but still have no
 * business in a form.
 */
const SYSTEM_COLUMN_NAMES = new Set([
  "ContentType", "Attachments", "Edit", "DocIcon", "ItemChildCount", "FolderChildCount", "AppAuthor",
  "AppEditor", "ComplianceAssetId",
]);

/** `LinkTitle`, `LinkTitleNoMenu`, `LinkTitle2`: the rendered-link twins of `Title`. */
const SYSTEM_COLUMN_PREFIXES = ["_", "LinkTitle"];

/** Internal column names are alphanumeric plus `_xHHHH_` escapes; nothing else may reach OData. */
const INTERNAL_NAME_PATTERN = /^[A-Za-z0-9_]+$/;

// ── Types ───────────────────────────────────────────────────────────
//
// The internal shapes this client speaks. Phase 2's agent-facing `sharepoint-types.d.ts` mirrors
// them; these are the source of truth for what a column can be and what a filter may say.

/**
 * What a column holds, from the mutually-exclusive facet Graph puts on a `columnDefinition`.
 *
 * `person` and `lookup` are recognized so they can be described and refused rather than silently
 * mistyped: both are written as `{name}LookupId` with a numeric id, which a form cannot supply.
 * `unsupported` covers every other facet (currency, calculated, hyperlink, term, geolocation, …):
 * readable as whatever Graph returns, never written.
 */
export type ColumnType =
  | "text" | "number" | "boolean" | "dateTime" | "choice" | "person" | "lookup" | "unsupported";

/** One column of a list, as a form or an agent needs to see it. */
export type ColumnDefinition = {
  /** Internal name. The one that appears in `fields`, `$select` and `$filter`. */
  name: string;
  /** Display name. Renameable, so never used to address anything. */
  displayName: string;
  type: ColumnType;
  required: boolean;
  readOnly: boolean;
  /** Allowed values, for `choice` columns. */
  choices?: string[];
  /** A `text` column that accepts more than one line. */
  multiline?: boolean;
  /** A `person` or `lookup` column that holds more than one value. */
  multiple?: boolean;
};

/** One list item: the listItem's own properties plus its expanded `fields`. */
export type ListItem = {
  id: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  fields: Record<string, unknown>;
};

/** Comparisons a caller may ask for. `startswith` is text-only, as in OData. */
export type WhereOperator = "eq" | "ne" | "lt" | "gt" | "le" | "ge" | "startswith";

/**
 * One filter clause, structured. There is deliberately no way to pass an OData string: the column is
 * matched against the list's real columns and the value is typed and escaped here.
 */
export type WhereClause = {
  column: string;
  op: WhereOperator;
  value: string | number | boolean;
};

/**
 * A site, or a list within one: the id to address it by, a name to show, and a link to open. Sites
 * and lists carry the same three, so they are described the same way.
 */
export type SharePointResourceInfo = {
  id: string;
  displayName: string;
  webUrl: string;
};

// ── Graph wire shapes ───────────────────────────────────────────────
// Only the fields this client reads. Everything is optional: these describe what Graph may send, not
// what it promises.

/** A site or a list, as Graph sends it. Both name themselves with the same fields. */
type GraphResource = {
  id?: string;
  displayName?: string;
  name?: string;
  webUrl?: string;
};

type GraphColumnDefinition = {
  name?: string;
  displayName?: string;
  hidden?: boolean;
  readOnly?: boolean;
  required?: boolean;
  text?: { allowMultipleLines?: boolean };
  number?: unknown;
  boolean?: unknown;
  dateTime?: unknown;
  choice?: { choices?: string[] };
  personOrGroup?: { allowMultipleSelection?: boolean };
  lookup?: { allowMultipleValues?: boolean };
};

type GraphListItem = {
  id?: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  fields?: Record<string, unknown>;
};

// ── Mapping ─────────────────────────────────────────────────────────

/**
 * Reduce a site or a list to what this client passes on. `noun` names the kind in the failure and in
 * the placeholder for an unnamed one, which is all the two cases differ by.
 */
function resourceInfoFrom(resource: GraphResource, noun: "site" | "list"): SharePointResourceInfo {
  if (typeof resource.id !== "string" || !resource.id) {
    throw new Error(`Microsoft Graph returned a SharePoint ${noun} without an id.`);
  }
  return {
    id: resource.id,
    displayName: resource.displayName || resource.name || `(unnamed ${noun})`,
    webUrl: typeof resource.webUrl === "string" ? resource.webUrl : "",
  };
}

/**
 * The last path component of a list's `webUrl`, decoded, or null when there is none.
 *
 * This is what a pasted URL's `/Lists/<name>` segment is matched against: the display name is
 * renameable while this URL name is fixed at creation, so it is the stable half of the pair.
 */
function listUrlName(webUrl: string | undefined): string | null {
  if (!webUrl) return null;
  let path: string;
  try {
    path = new URL(webUrl).pathname;
  } catch {
    return null;
  }
  let last = path.split("/").filter(segment => segment !== "").pop();
  if (!last) return null;
  try {
    return decodeURIComponent(last);
  } catch {
    return last;
  }
}

/** Whether a column is plumbing rather than data. */
function isSystemColumn(name: string): boolean {
  return SYSTEM_COLUMN_NAMES.has(name) ||
      SYSTEM_COLUMN_PREFIXES.some(prefix => name.startsWith(prefix));
}

function columnTypeOf(column: GraphColumnDefinition): ColumnType {
  // Facets are mutually exclusive on a columnDefinition, so the first match is the answer. Anything
  // outside this set (currency, calculated, hyperlink, term, …) is readable but not writable, which
  // is what `unsupported` means to the caller.
  if (column.text) return "text";
  if (column.number) return "number";
  if (column.boolean) return "boolean";
  if (column.dateTime) return "dateTime";
  if (column.choice) return "choice";
  if (column.personOrGroup) return "person";
  if (column.lookup) return "lookup";
  return "unsupported";
}

/**
 * Normalise one Graph `columnDefinition`, or null when it is not a column a caller should see.
 *
 * Hidden and read-only columns are dropped: neither can be filled in, and the read-only set is where
 * SharePoint keeps its own bookkeeping (`ID`, `Created`, `Author`, …). `Title` survives the drop
 * unconditionally — it is the one column every list has, it is what a list's items are named by, and
 * some lists mark it read-only in a view definition while still accepting it on create.
 */
function normalizeColumn(column: GraphColumnDefinition): ColumnDefinition | null {
  let name = typeof column.name === "string" ? column.name : "";
  if (!name || !INTERNAL_NAME_PATTERN.test(name)) return null;

  let isTitle = name === "Title";
  if (!isTitle && (column.hidden === true || column.readOnly === true)) return null;
  if (!isTitle && isSystemColumn(name)) return null;

  let type = columnTypeOf(column);
  let choices = column.choice?.choices;
  let multiple = column.personOrGroup?.allowMultipleSelection ?? column.lookup?.allowMultipleValues;
  return {
    name,
    displayName: typeof column.displayName === "string" && column.displayName
        ? column.displayName
        : name,
    type,
    required: column.required === true,
    readOnly: column.readOnly === true,
    ...(type === "choice" ? { choices: Array.isArray(choices) ? choices.filter(
        (choice): choice is string => typeof choice === "string") : [] } : {}),
    ...(type === "text" && column.text?.allowMultipleLines === true ? { multiline: true } : {}),
    ...(multiple === true ? { multiple: true } : {}),
  };
}

function listItemFrom(item: GraphListItem): ListItem {
  if (typeof item.id !== "string" || !item.id) {
    throw new Error("Microsoft Graph returned a list item without an id.");
  }
  let fields: Record<string, unknown> = {};
  for (let [key, value] of Object.entries(item.fields ?? {})) {
    // `@odata.etag` and friends are OData control information, not list data.
    if (!key.startsWith("@")) fields[key] = value;
  }
  return {
    id: item.id,
    ...(typeof item.webUrl === "string" ? { webUrl: item.webUrl } : {}),
    ...(typeof item.createdDateTime === "string"
        ? { createdDateTime: item.createdDateTime }
        : {}),
    ...(typeof item.lastModifiedDateTime === "string"
        ? { lastModifiedDateTime: item.lastModifiedDateTime }
        : {}),
    fields,
  };
}

// ── Filtering ───────────────────────────────────────────────────────

/** An OData string literal: single-quoted, with any quote inside doubled, as OData escapes it. */
function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

function describeValue(value: string | number | boolean): string {
  return typeof value === "string" ? "text" : typeof value;
}

/** The literal for one clause's value, type-checked against the column it is compared with. */
function filterLiteral(column: ColumnDefinition, value: string | number | boolean): string {
  let mismatch = () => new Error(
      `Column "${column.name}" holds ${column.type} values, so it cannot be compared with ` +
      `${describeValue(value)}.`);

  switch (column.type) {
    case "text":
    case "choice": {
      if (typeof value !== "string") throw mismatch();
      if (value.length > MAX_FILTER_TEXT_CHARS) {
        throw new Error(
            `A filter value must be at most ${MAX_FILTER_TEXT_CHARS} characters.`);
      }
      return odataString(value);
    }
    case "number": {
      if (typeof value !== "number" || !Number.isFinite(value)) throw mismatch();
      return String(value);
    }
    case "boolean": {
      if (typeof value !== "boolean") throw mismatch();
      return value ? "true" : "false";
    }
    case "dateTime": {
      if (typeof value !== "string") throw mismatch();
      let parsed = Date.parse(value);
      if (Number.isNaN(parsed)) {
        throw new Error(
            `Column "${column.name}" holds dates, so its value must be an ISO 8601 date or ` +
            "date-time.");
      }
      // Re-emitted from the parsed date rather than passed through: the literal is then known to be
      // a plain ISO timestamp, whatever the caller wrote.
      return odataString(new Date(parsed).toISOString());
    }
    default:
      throw new Error(
          `Column "${column.name}" is a ${column.type} column, which cannot be filtered on.`);
  }
}

/**
 * Build the `$filter` for a list-items read from clauses validated against the list's own columns.
 *
 * Every part of the result comes from this file or from Graph: the column name is one of the list's
 * real internal names, the operator is one of seven literals, and the value is a typed literal with
 * OData escaping applied. A caller therefore cannot express an OData expression — which is the whole
 * point, since the caller is an agent.
 *
 * Returns the empty string for no clauses, so a caller can drop the parameter.
 */
export function buildItemsFilter(where: WhereClause[], columns: ColumnDefinition[]): string {
  if (where.length === 0) return "";
  if (where.length > MAX_WHERE_CLAUSES) {
    throw new Error(
        `A filter may combine at most ${MAX_WHERE_CLAUSES} conditions; ${where.length} were given.`);
  }

  return where.map(clause => {
    let column = columns.find(candidate => candidate.name === clause.column);
    if (!column) {
      throw new Error(
          `Unknown column "${truncate(String(clause.column), MAX_NAME_ECHO_CHARS)}". Use an ` +
          "internal column name from getColumns().");
    }
    // Columns arrive from Graph, so this is a guard against a hostile response rather than a hostile
    // caller — but it is what makes "nothing unescaped reaches OData" true without qualification.
    if (!INTERNAL_NAME_PATTERN.test(column.name)) {
      throw new Error(`Column "${truncate(column.name, MAX_NAME_ECHO_CHARS)}" cannot be filtered.`);
    }

    let literal = filterLiteral(column, clause.value);
    if (clause.op === "startswith") {
      if (column.type !== "text") {
        throw new Error(
            `"startswith" only applies to text columns, and "${column.name}" is a ${column.type} ` +
            "column.");
      }
      return `startswith(fields/${column.name},${literal})`;
    }
    if (!["eq", "ne", "lt", "gt", "le", "ge"].includes(clause.op)) {
      throw new Error(
          `Unknown comparison "${truncate(String(clause.op), MAX_NAME_ECHO_CHARS)}".`);
    }
    return `fields/${column.name} ${clause.op} ${literal}`;
  }).join(" and ");
}

// ── Client ──────────────────────────────────────────────────────────

/** The throttling ceiling, as an error the caller can act on. Shared with the apply path. */
export function sharePointThrottled(requestedMs: number): GraphApiError {
  return new GraphApiError(429, "activityLimitReached",
      "SharePoint is throttling this connection and asked to be left alone for " +
      `${Math.round(requestedMs / 1000)} seconds. Try again after that.`);
}

export type GraphSharePointApiOptions = {
  /** Called when Graph answers a claims challenge, so the account can be marked dead. */
  onCredentialsRejected?: CredentialsRejectedReporter;
  /**
   * How long a throttling `Retry-After` is waited out before the request fails fast. Defaults to the
   * read ceiling; the apply path passes a longer one.
   */
  retryAfter?: RetryAfterPolicy;
};

export type ListItemsOptions = {
  /**
   * The list's columns, as `listColumns` returned them. Selection and filtering are validated
   * against these, so a name that is not a real column never reaches Graph.
   */
  columns: ColumnDefinition[];
  /** Internal column names to read back. Defaults to every column in `columns`. */
  select?: string[];
  /** Structured filter clauses. Defaults to unfiltered. */
  where?: WhereClause[];
  /** Items per page. Clamped to `MAX_ITEM_PAGE_SIZE`. */
  top?: number;
};

export class GraphSharePointApi {
  #getAccessToken: AccessTokenProvider;
  #onCredentialsRejected: CredentialsRejectedReporter | undefined;
  #retryAfter: RetryAfterPolicy;

  constructor(getAccessToken: AccessTokenProvider, opts: GraphSharePointApiOptions = {}) {
    this.#getAccessToken = getAccessToken;
    this.#onCredentialsRejected = opts.onCredentialsRejected;
    this.#retryAfter = opts.retryAfter ??
        { maxWaitMs: DEFAULT_MAX_RETRY_AFTER_MS, tooLong: sharePointThrottled };
  }

  /**
   * The one place a SharePoint Graph request is made.
   *
   * The decorated `User-Agent` and the throttling ceiling are set here, unconditionally, so no call
   * site can make a request without them. No `Prefer` header is sent: see the file header for the
   * three that might have been.
   */
  async #request(url: string, init: RequestInit = {}): Promise<Response> {
    let headers = new Headers(init.headers);
    headers.set("User-Agent", SHAREPOINT_USER_AGENT);
    if (!headers.has("Accept")) headers.set("Accept", "application/json");
    if (init.body && !headers.has("Content-Type")) headers.set("Content-Type", "application/json");

    return await fetchWithAuthRetry(url, { ...init, headers }, this.#getAccessToken, {
      timeoutMs: GRAPH_TIMEOUT_MS,
      retryAfter: this.#retryAfter,
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

    if (response.status === 400) {
      // Graph's own words, kept verbatim. This is where "Field 'X' cannot be referenced in filter
      // or orderby as it is not indexed" arrives, and that sentence names the column and the fix.
      return new GraphApiError(400, code || "invalidRequest",
          `SharePoint rejected the request: ${detail}`);
    }

    if (response.status === 401) {
      let claims = claimsChallengeDetail(response.headers.get("WWW-Authenticate"));
      if (claims) {
        // fetchWithAuthRetry has already reported this to the account, which is what makes the
        // Workshop offer a reconnect. Nothing here retries.
        return new GraphApiError(401, code || claims,
            "Microsoft requires this account to sign in again before SharePoint can be used " +
            `(${claims}). Please reconnect the account.`,
            { credentialsRejected: true });
      }
      return new GraphApiError(401, code || "unauthenticated",
          `Microsoft rejected the SharePoint credentials (${detail}).`);
    }

    if (response.status === 403) {
      return new GraphApiError(403, code || "forbidden",
          `This Microsoft connection is not permitted to use that SharePoint list (${detail}). ` +
          "Check that the connected account can open the list, and that an administrator has " +
          "granted the SharePoint permissions.");
    }

    if (response.status === 404) {
      return new GraphApiError(404, code || "notFound",
          `That SharePoint site, list or item no longer exists, or this account cannot see it ` +
          `(${detail}).`);
    }

    if (response.status === 429) {
      let retryAfter = response.headers.get("Retry-After");
      return new GraphApiError(429, code || "activityLimitReached",
          "SharePoint is throttling this connection" +
          (retryAfter ? `; retry in ${truncate(retryAfter, 20)} seconds` : "") + ".");
    }

    if (response.status >= 500) {
      return new GraphApiError(response.status, code || "serviceError",
          `SharePoint is temporarily unavailable (${detail}).`);
    }

    return new GraphApiError(response.status, code || "unknown",
        `Microsoft Graph request failed: ${response.status}${code ? ` ${code}` : ""} — ${detail}`);
  }

  // ── Resolution ────────────────────────────────────────────────────────────────

  /**
   * The site a pasted URL's host and path name.
   *
   * `sitePath` is the decoded server-relative path from `parseSharePointListUrl`; empty resolves the
   * tenant root site.
   */
  async resolveSite(hostname: string, sitePath: string): Promise<SharePointResourceInfo> {
    return resourceInfoFrom(await this.#fetchJson<GraphResource>(
        graphSiteByPathUrl(hostname, sitePath, { $select: "id,displayName,webUrl" })), "site");
  }

  /**
   * The list on a site whose URL name matches a pasted URL's `/Lists/<name>` segment.
   *
   * Graph has no "get list by url name" endpoint, so the site's lists are walked and matched on the
   * last segment of each `webUrl`, case-insensitively — SharePoint URLs are case-preserving but not
   * case-sensitive, so a pasted link can differ in case from what Graph reports.
   */
  async resolveListByUrl(siteId: string, listSegment: string): Promise<SharePointResourceInfo> {
    let wanted = listSegment.toLowerCase();
    let url: string = graphUrl(["sites", siteId, "lists"], {
      $select: "id,displayName,webUrl",
      $top: String(LIST_PAGE_SIZE),
    });

    for (let page = 0; page < MAX_LIST_PAGES; page++) {
      let body = await this.#fetchJson<GraphCollection<GraphResource>>(url);
      for (let list of body.value ?? []) {
        if (listUrlName(list.webUrl)?.toLowerCase() === wanted) {
          return resourceInfoFrom(list, "list");
        }
      }
      let next = body["@odata.nextLink"];
      if (!next) {
        throw new GraphApiError(404, "listNotFound",
            "That list was not found on the site, or this account cannot open it. Check the URL " +
            "and that the list is shared with you.");
      }
      url = assertGraphUrl(next);
    }

    throw new GraphApiError(404, "listNotFound",
        `That site has more lists than this can search (${MAX_LIST_PAGES} pages). Open the list ` +
        "in SharePoint and paste its URL from there.");
  }

  /** The list's current name and link. Used to keep a binding's title honest after a rename. */
  async describeList(siteId: string, listId: string): Promise<SharePointResourceInfo> {
    let list = await this.#fetchJson<GraphResource>(
        graphUrl(["sites", siteId, "lists", listId], { $select: "id,displayName,webUrl" }));
    return resourceInfoFrom(list, "list");
  }

  /** The list's columns, normalised, with SharePoint's plumbing dropped. */
  async listColumns(siteId: string, listId: string): Promise<ColumnDefinition[]> {
    let body = await this.#fetchJson<GraphCollection<GraphColumnDefinition>>(
        graphUrl(["sites", siteId, "lists", listId, "columns"]));
    return (body.value ?? [])
        .map(normalizeColumn)
        .filter((column): column is ColumnDefinition => column !== null);
  }

  // ── Items ─────────────────────────────────────────────────────────────────────

  /**
   * First page of items. `select` and `where` are validated against `columns`, so neither a column
   * name nor a filter fragment can be invented by the caller.
   */
  async listItems(siteId: string, listId: string, opts: ListItemsOptions)
      : Promise<GraphPage<ListItem>> {
    let selected = this.#selectedFields(opts);
    let filter = buildItemsFilter(opts.where ?? [], opts.columns);
    return await this.#itemPage(graphUrl(["sites", siteId, "lists", listId, "items"], {
      // `fields` alone returns every column, which is the right answer for a list whose columns are
      // all plumbing — there is nothing to narrow to.
      $expand: selected.length > 0 ? `fields($select=${selected.join(",")})` : "fields",
      $top: itemPageSize(opts.top),
      ...(filter ? { $filter: filter } : {}),
    }));
  }

  /** Follow a `@odata.nextLink` produced by a previous page. */
  async nextItemPage(nextLink: string): Promise<GraphPage<ListItem>> {
    return await this.#itemPage(assertGraphUrl(nextLink));
  }

  /** One item, with every field expanded. */
  async getItem(siteId: string, listId: string, itemId: string): Promise<ListItem> {
    return listItemFrom(await this.#fetchJson<GraphListItem>(
        graphUrl(["sites", siteId, "lists", listId, "items", itemId], { $expand: "fields" })));
  }

  /**
   * Create an item.
   *
   * `fields` is validated by the caller against the list's columns before it gets here — this client
   * has no opinion on what a list's columns mean, only on how they are addressed.
   */
  async createItem(siteId: string, listId: string, fields: Record<string, unknown>)
      : Promise<ListItem> {
    return listItemFrom(await this.#fetchJson<GraphListItem>(
        graphUrl(["sites", siteId, "lists", listId, "items"]), {
          method: "POST",
          body: JSON.stringify({ fields }),
        }));
  }

  /** The internal names to expand, defaulting to every column the list has. */
  #selectedFields(opts: ListItemsOptions): string[] {
    if (!opts.select) return opts.columns.map(column => column.name);
    return opts.select.map(name => {
      let column = opts.columns.find(candidate => candidate.name === name);
      if (!column) {
        throw new Error(
            `Unknown column "${truncate(String(name), MAX_NAME_ECHO_CHARS)}". Use internal column ` +
            "names from getColumns().");
      }
      return column.name;
    });
  }

  async #itemPage(url: string): Promise<GraphPage<ListItem>> {
    let body = await this.#fetchJson<GraphCollection<GraphListItem>>(url);
    let next = body["@odata.nextLink"];
    return {
      items: (body.value ?? []).map(listItemFrom),
      // Validated on arrival, not only when followed, so a bad link fails the page that produced it.
      ...(next ? { nextLink: assertGraphUrl(next) } : {}),
    };
  }
}

/** `$top` as Graph wants it: a positive integer within this client's ceiling. */
function itemPageSize(requested: number | undefined): string {
  let size = requested ?? DEFAULT_ITEM_PAGE_SIZE;
  if (!Number.isFinite(size)) size = DEFAULT_ITEM_PAGE_SIZE;
  return String(Math.min(Math.max(Math.trunc(size), 1), MAX_ITEM_PAGE_SIZE));
}
