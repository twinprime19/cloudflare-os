import { afterEach, describe, expect, it, vi } from "vitest";

import { GraphApiError, graphSiteByPathUrl } from "../src/graph-api";
import {
  ColumnDefinition, GraphSharePointApi, SHAREPOINT_USER_AGENT, buildItemsFilter,
} from "../src/graph-sharepoint-api";

type Call = { url: string; init: RequestInit };

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status, headers: { "Content-Type": "application/json", ...headers },
  });
}

/** Route every fetch through `handler`, recording what the client sent. */
function stubFetch(handler: (call: Call) => Response | Promise<Response>): Call[] {
  const calls: Call[] = [];
  vi.stubGlobal("fetch", vi.fn(async (url: string, init: RequestInit = {}) => {
    calls.push({ url, init });
    return await handler({ url, init });
  }));
  return calls;
}

function header(call: Call, name: string): string | null {
  return new Headers(call.init.headers).get(name);
}

function param(call: Call, name: string): string | null {
  return new URL(call.url).searchParams.get(name);
}

function newApi(opts: ConstructorParameters<typeof GraphSharePointApi>[1] = {}) {
  return new GraphSharePointApi(async () => "access-token", opts);
}

/** A list with one column of every type this client recognizes, plus SharePoint's plumbing. */
const GRAPH_COLUMNS = [
  { name: "Title", displayName: "Title", required: true, text: { allowMultipleLines: false } },
  { name: "Details", displayName: "Details", text: { allowMultipleLines: true } },
  { name: "Qty", displayName: "Quantity", number: { decimalPlaces: "none" } },
  { name: "Done", displayName: "Done", boolean: {} },
  { name: "Due", displayName: "Due date", dateTime: { format: "dateOnly" } },
  { name: "Status", displayName: "Status", required: true, choice: { choices: ["New", "Open"] } },
  { name: "Owner", displayName: "Owner", personOrGroup: { allowMultipleSelection: true } },
  { name: "Ref", displayName: "Reference", lookup: { allowMultipleValues: false } },
  { name: "Budget", displayName: "Budget", currency: { locale: "en-US" } },
  { name: "ID", displayName: "ID", readOnly: true, number: {} },
  { name: "Created", displayName: "Created", readOnly: true, dateTime: {} },
  { name: "Secret", displayName: "Secret", hidden: true, text: {} },
  { name: "_UIVersionString", displayName: "Version", text: {} },
  { name: "LinkTitleNoMenu", displayName: "Title", text: {} },
  { name: "ContentType", displayName: "Content Type", text: {} },
  { name: "Attachments", displayName: "Attachments", boolean: {} },
];

/** The same list, already normalised — what a caller passes back in when reading or filtering. */
const SCHEMA: ColumnDefinition[] = [
  { name: "Title", displayName: "Title", type: "text", required: true, readOnly: false },
  { name: "Qty", displayName: "Quantity", type: "number", required: false, readOnly: false },
  { name: "Done", displayName: "Done", type: "boolean", required: false, readOnly: false },
  { name: "Due", displayName: "Due date", type: "dateTime", required: false, readOnly: false },
  {
    name: "Status", displayName: "Status", type: "choice", required: true, readOnly: false,
    choices: ["New", "Open"],
  },
  { name: "Owner", displayName: "Owner", type: "person", required: false, readOnly: false },
  { name: "Ref", displayName: "Reference", type: "lookup", required: false, readOnly: false },
];

const ITEM = {
  id: "42",
  webUrl: "https://contoso.sharepoint.com/sites/HR/Lists/Requests/42_.000",
  createdDateTime: "2026-09-15T02:00:00Z",
  lastModifiedDateTime: "2026-09-15T02:00:00Z",
  fields: { "@odata.etag": "\"1\"", Title: "New laptop", Qty: 2 },
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("site addressing", () => {
  it("addresses a site by host and server-relative path", () => {
    expect(graphSiteByPathUrl("contoso.sharepoint.com", "/sites/HR"))
      .toBe("https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/HR");
    expect(graphSiteByPathUrl("contoso.sharepoint.com", "/sites/HR/sub"))
      .toBe("https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/HR/sub");
    expect(graphSiteByPathUrl("contoso.sharepoint.com", "/teams/HR"))
      .toBe("https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/teams/HR");
  });

  it("addresses the tenant root site by host alone", () => {
    expect(graphSiteByPathUrl("contoso.sharepoint.com", ""))
      .toBe("https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com");
  });

  it("percent-encodes each path component instead of letting it become syntax", () => {
    expect(graphSiteByPathUrl("contoso.sharepoint.com", "/sites/H R", { $select: "id" }))
      .toBe("https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/H%20R" +
            "?$select=id");
  });

  it("refuses a relative or empty component, which URL parsing would collapse", () => {
    // `/sites/contoso.sharepoint.com:/sites/../secret` would resolve to a different site, reached
    // with this account's bearer token.
    expect(() => graphSiteByPathUrl("contoso.sharepoint.com", "/sites/../secret"))
      .toThrow(/empty or relative/);
    expect(() => graphSiteByPathUrl("contoso.sharepoint.com", "/sites//secret"))
      .toThrow(/empty or relative/);
    expect(() => graphSiteByPathUrl("", "/sites/HR")).toThrow(/empty or relative/);
  });

  it("resolves a site through that address, selecting only what a binding needs", async () => {
    const calls = stubFetch(() => jsonResponse({
      id: "site-1", displayName: "IT Department", webUrl: "https://contoso.sharepoint.com/sites/IT",
    }));

    const site = await newApi().resolveSite("contoso.sharepoint.com", "/sites/IT");

    expect(site).toEqual({
      id: "site-1", displayName: "IT Department", webUrl: "https://contoso.sharepoint.com/sites/IT",
    });
    expect(calls[0].url).toBe(
      "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/IT" +
      `?$select=${encodeURIComponent("id,displayName,webUrl")}`);
  });
});

describe("request headers", () => {
  it("decorates every request with the SharePoint user agent", async () => {
    const calls = stubFetch(call => call.url.includes("/items/42") || call.init.method === "POST"
      ? jsonResponse(ITEM, call.init.method === "POST" ? 201 : 200)
      : jsonResponse({ value: [ITEM] }));
    const api = newApi();

    await api.listItems("site-1", "list-1", { columns: SCHEMA });
    await api.getItem("site-1", "list-1", "42");
    await api.createItem("site-1", "list-1", { Title: "New laptop" });

    expect(calls).toHaveLength(3);
    for (const call of calls) {
      expect(header(call, "User-Agent")).toBe(SHAREPOINT_USER_AGENT);
      expect(header(call, "Accept")).toBe("application/json");
      // The Outlook preferences are mail-only; SharePoint gets none of them.
      expect(header(call, "Prefer")).toBeNull();
    }
  });
});

/** A `/lists` page whose entries are addressed by URL name, as SharePoint reports them. */
function lists(...names: string[]) {
  return {
    value: names.map((name, index) => ({
      id: `list-${index}`,
      displayName: name.replace(/%20/g, " "),
      webUrl: `https://contoso.sharepoint.com/sites/HR/Lists/${name}`,
    })),
  };
}

describe("resolveListByUrl", () => {
  it("matches the last segment of a list's webUrl, decoded and case-insensitively", async () => {
    const calls = stubFetch(() => jsonResponse(lists("Other", "DEV%20Contoso%20OS%20Issues%20Log")));

    const list = await newApi().resolveListByUrl("site-1", "dev contoso os issues log");

    expect(list.id).toBe("list-1");
    expect(list.displayName).toBe("DEV Contoso OS Issues Log");
    expect(calls[0].url).toBe(
      "https://graph.microsoft.com/v1.0/sites/site-1/lists" +
      `?$select=${encodeURIComponent("id,displayName,webUrl")}&$top=100`);
  });

  it("follows the next link when the first page has no match", async () => {
    const next = "https://graph.microsoft.com/v1.0/sites/site-1/lists?$skiptoken=abc";
    const calls = stubFetch(call => call.url === next
      ? jsonResponse(lists("Requests"))
      : jsonResponse({ ...lists("Other"), "@odata.nextLink": next }));

    const list = await newApi().resolveListByUrl("site-1", "Requests");

    expect(list.id).toBe("list-0");
    expect(calls).toHaveLength(2);
    expect(calls[1].url).toBe(next);
  });

  it("refuses a next link outside Graph, which would carry the bearer token off-origin", async () => {
    stubFetch(() => jsonResponse({
      ...lists("Other"), "@odata.nextLink": "https://evil.example/v1.0/sites/site-1/lists",
    }));

    await expect(newApi().resolveListByUrl("site-1", "Requests"))
      .rejects.toThrow(/outside https:\/\/graph\.microsoft\.com/);
  });

  it("says the list was not found when the site has no match", async () => {
    stubFetch(() => jsonResponse(lists("Other")));

    await expect(newApi().resolveListByUrl("site-1", "Requests"))
      .rejects.toThrow(/was not found on the site, or this account cannot open it/);
  });

  it("stops after a bounded number of pages rather than walking a site forever", async () => {
    const calls = stubFetch(() => jsonResponse({
      ...lists("Other"),
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites/site-1/lists?$skiptoken=abc",
    }));

    await expect(newApi().resolveListByUrl("site-1", "Requests"))
      .rejects.toThrow(/more lists than this can search/);
    expect(calls).toHaveLength(10);
  });
});

describe("listColumns", () => {
  it("normalises each facet Graph reports onto a type a form can use", async () => {
    const calls = stubFetch(() => jsonResponse({ value: GRAPH_COLUMNS }));

    const columns = await newApi().listColumns("site-1", "list-1");

    expect(calls[0].url).toBe("https://graph.microsoft.com/v1.0/sites/site-1/lists/list-1/columns");
    expect(columns).toEqual([
      { name: "Title", displayName: "Title", type: "text", required: true, readOnly: false },
      {
        name: "Details", displayName: "Details", type: "text", required: false, readOnly: false,
        multiline: true,
      },
      { name: "Qty", displayName: "Quantity", type: "number", required: false, readOnly: false },
      { name: "Done", displayName: "Done", type: "boolean", required: false, readOnly: false },
      { name: "Due", displayName: "Due date", type: "dateTime", required: false, readOnly: false },
      {
        name: "Status", displayName: "Status", type: "choice", required: true, readOnly: false,
        choices: ["New", "Open"],
      },
      {
        name: "Owner", displayName: "Owner", type: "person", required: false, readOnly: false,
        multiple: true,
      },
      { name: "Ref", displayName: "Reference", type: "lookup", required: false, readOnly: false },
      // Currency has no write shape this gatekeeper knows, so it is described but not writable.
      { name: "Budget", displayName: "Budget", type: "unsupported", required: false,
        readOnly: false },
    ]);
  });

  it("keeps Title even when the list marks it read-only", async () => {
    stubFetch(() => jsonResponse({
      value: [{ name: "Title", displayName: "Title", readOnly: true, text: {} }],
    }));

    const columns = await newApi().listColumns("site-1", "list-1");

    expect(columns).toEqual([
      { name: "Title", displayName: "Title", type: "text", required: false, readOnly: true },
    ]);
  });
});

describe("listItems", () => {
  it("expands only the list's own columns and clamps the page size", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [ITEM] }));

    const page = await newApi().listItems("site-1", "list-1", {
      columns: SCHEMA, select: ["Title", "Qty"], top: 5000,
    });

    expect(new URL(calls[0].url).pathname)
      .toBe("/v1.0/sites/site-1/lists/list-1/items");
    expect(param(calls[0], "$expand")).toBe("fields($select=Title,Qty)");
    expect(param(calls[0], "$top")).toBe("200");
    expect(param(calls[0], "$filter")).toBeNull();
    // OData control information is not list data, so it never reaches the caller.
    expect(page.items).toEqual([{
      id: "42",
      webUrl: ITEM.webUrl,
      createdDateTime: ITEM.createdDateTime,
      lastModifiedDateTime: ITEM.lastModifiedDateTime,
      fields: { Title: "New laptop", Qty: 2 },
    }]);
  });

  it("selects every column by default and uses the default page size", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [] }));

    await newApi().listItems("site-1", "list-1", { columns: SCHEMA });

    expect(param(calls[0], "$expand"))
      .toBe("fields($select=Title,Qty,Done,Due,Status,Owner,Ref)");
    expect(param(calls[0], "$top")).toBe("25");
  });

  it("refuses a column name that is not on the list, before any request is made", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [] }));

    await expect(newApi().listItems("site-1", "list-1", {
      columns: SCHEMA, select: ["Title", "Password"],
    })).rejects.toThrow(/Unknown column "Password"/);

    expect(calls).toHaveLength(0);
  });

  it("builds the filter from clauses rather than accepting an OData string", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [] }));

    await newApi().listItems("site-1", "list-1", {
      columns: SCHEMA, where: [{ column: "Status", op: "eq", value: "O'Brien" }],
    });

    expect(param(calls[0], "$filter")).toBe("fields/Status eq 'O''Brien'");
  });

  it("validates a next link when the page that carried it arrives", async () => {
    stubFetch(() => jsonResponse({
      value: [], "@odata.nextLink": "https://evil.example/v1.0/sites/site-1/lists/list-1/items",
    }));

    await expect(newApi().listItems("site-1", "list-1", { columns: SCHEMA }))
      .rejects.toThrow(/outside https:\/\/graph\.microsoft\.com/);
  });

  it("pins a next link it is asked to follow", async () => {
    const calls = stubFetch(() => jsonResponse({ value: [ITEM] }));

    await expect(newApi().nextItemPage("https://evil.example/v1.0/items"))
      .rejects.toThrow(/outside https:\/\/graph\.microsoft\.com/);
    expect(calls).toHaveLength(0);

    const page = await newApi().nextItemPage(
      "https://graph.microsoft.com/v1.0/sites/site-1/lists/list-1/items?$skiptoken=abc");
    expect(page.items[0].id).toBe("42");
  });
});

describe("getItem and createItem", () => {
  it("reads one item with its fields expanded", async () => {
    const calls = stubFetch(() => jsonResponse(ITEM));

    const item = await newApi().getItem("site-1", "list-1", "42");

    expect(calls[0].url).toBe(
      "https://graph.microsoft.com/v1.0/sites/site-1/lists/list-1/items/42?$expand=fields");
    expect(item.id).toBe("42");
    expect(item.fields).toEqual({ Title: "New laptop", Qty: 2 });
  });

  it("keeps an item id inside one path segment", async () => {
    const calls = stubFetch(() => jsonResponse(ITEM));

    await expect(newApi().getItem("site-1", "list-1", "..")).rejects.toThrow(/empty or relative/);
    expect(calls).toHaveLength(0);

    await newApi().getItem("site-1", "list-1", "../../drives/victim");
    expect(new URL(calls[0].url).pathname.split("/")).toHaveLength(8);
  });

  it("posts the fields in the envelope Graph expects and returns the created item", async () => {
    const calls = stubFetch(() => jsonResponse({ ...ITEM, id: "43" }, 201));

    const created = await newApi().createItem("site-1", "list-1", {
      Title: "New laptop", Qty: 2,
    });

    expect(calls[0].url).toBe("https://graph.microsoft.com/v1.0/sites/site-1/lists/list-1/items");
    expect(calls[0].init.method).toBe("POST");
    expect(header(calls[0], "Content-Type")).toBe("application/json");
    expect(JSON.parse(String(calls[0].init.body)))
      .toEqual({ fields: { Title: "New laptop", Qty: 2 } });
    expect(created.id).toBe("43");
  });
});

describe("buildItemsFilter", () => {
  it("emits nothing for no clauses", () => {
    expect(buildItemsFilter([], SCHEMA)).toBe("");
  });

  it("types each value against the column it is compared with", () => {
    expect(buildItemsFilter([
      { column: "Qty", op: "lt", value: 600 },
      { column: "Done", op: "eq", value: false },
      { column: "Due", op: "ge", value: "2026-01-01T00:00:00Z" },
      { column: "Title", op: "startswith", value: "Lap" },
    ], SCHEMA)).toBe(
      "fields/Qty lt 600 and fields/Done eq false and " +
      "fields/Due ge '2026-01-01T00:00:00.000Z' and startswith(fields/Title,'Lap')");
  });

  it("doubles a quote so a value cannot close its own literal", () => {
    expect(buildItemsFilter([{ column: "Title", op: "eq", value: "' or startswith(fields/x,'" }],
                            SCHEMA))
      .toBe("fields/Title eq ''' or startswith(fields/x,'''");
  });

  it("refuses a column the list does not have", () => {
    expect(() => buildItemsFilter([{ column: "fields/Title eq 'x' or Password", op: "eq", value: "" }],
                                  SCHEMA))
      .toThrow(/Unknown column/);
  });

  it("refuses a value whose type does not match the column", () => {
    expect(() => buildItemsFilter([{ column: "Qty", op: "eq", value: "600" }], SCHEMA))
      .toThrow(/holds number values, so it cannot be compared with text/);
    expect(() => buildItemsFilter([{ column: "Title", op: "eq", value: 600 }], SCHEMA))
      .toThrow(/holds text values, so it cannot be compared with number/);
    expect(() => buildItemsFilter([{ column: "Due", op: "eq", value: "whenever" }], SCHEMA))
      .toThrow(/must be an ISO 8601 date/);
  });

  it("refuses columns whose OData shape this client does not write", () => {
    expect(() => buildItemsFilter([{ column: "Owner", op: "eq", value: "bob" }], SCHEMA))
      .toThrow(/is a person column, which cannot be filtered on/);
    expect(() => buildItemsFilter([{ column: "Ref", op: "eq", value: "1" }], SCHEMA))
      .toThrow(/is a lookup column, which cannot be filtered on/);
  });

  it("allows startswith on text only", () => {
    expect(() => buildItemsFilter([{ column: "Status", op: "startswith", value: "New" }], SCHEMA))
      .toThrow(/"startswith" only applies to text columns/);
  });

  it("refuses an unknown comparison", () => {
    expect(() => buildItemsFilter(
      [{ column: "Title", op: "contains" as "eq", value: "x" }], SCHEMA))
      .toThrow(/Unknown comparison "contains"/);
  });

  it("caps the number of clauses and the length of a value", () => {
    const clause = { column: "Title", op: "eq", value: "x" } as const;
    expect(() => buildItemsFilter([clause, clause, clause, clause, clause, clause], SCHEMA))
      .toThrow(/at most 5 conditions/);
    expect(() => buildItemsFilter([{ column: "Title", op: "eq", value: "x".repeat(257) }], SCHEMA))
      .toThrow(/at most 256 characters/);
  });
});

describe("error mapping", () => {
  it("passes Graph's non-indexed-column explanation through verbatim", async () => {
    const detail = "Field 'Qty' cannot be referenced in filter or orderby as it is not indexed. " +
      "Consider using indexed fields or sorting.";
    stubFetch(() => jsonResponse({ error: { code: "invalidRequest", message: detail } }, 400));

    const failure = await newApi().listItems("site-1", "list-1", {
      columns: SCHEMA, where: [{ column: "Qty", op: "lt", value: 600 }],
    }).catch((error: unknown) => error as GraphApiError);

    expect(failure).toBeInstanceOf(GraphApiError);
    expect(failure.status).toBe(400);
    expect(failure.message).toContain(detail);
  });

  it("reports a claims challenge as credentials the user must reconnect", async () => {
    const rejected = vi.fn(async () => {});
    stubFetch(() => jsonResponse({ error: { code: "InvalidAuthenticationToken" } }, 401, {
      "WWW-Authenticate": "Bearer error=\"insufficient_claims\", claims=\"eyJhbGci\"",
    }));

    const failure = await newApi({ onCredentialsRejected: rejected })
      .describeList("site-1", "list-1")
      .catch((error: unknown) => error as GraphApiError);

    expect(rejected).toHaveBeenCalledOnce();
    expect(failure.credentialsRejected).toBe(true);
    expect(failure.message).toMatch(/sign in again before SharePoint can be used/);
  });

  it("explains a 403 as a permission or sharing problem", async () => {
    stubFetch(() => jsonResponse({ error: { code: "accessDenied", message: "denied" } }, 403));

    await expect(newApi().describeList("site-1", "list-1"))
      .rejects.toThrow(/not permitted to use that SharePoint list/);
  });

  it("explains a 404 as gone or invisible to this account", async () => {
    stubFetch(() => jsonResponse({ error: { code: "itemNotFound", message: "gone" } }, 404));

    await expect(newApi().getItem("site-1", "list-1", "42"))
      .rejects.toThrow(/no longer exists, or this account cannot see it/);
  });
});

describe("throttling", () => {
  it("fails fast instead of sitting out a wait longer than the read ceiling", async () => {
    const calls = stubFetch(() => jsonResponse({ error: { code: "activityLimitReached" } }, 429, {
      "Retry-After": "120",
    }));

    const failure = await newApi().listItems("site-1", "list-1", { columns: SCHEMA })
      .catch((error: unknown) => error as GraphApiError);

    // One attempt, then the wait is reported rather than served: replaying early would land back in
    // the same throttle and spend another request on it.
    expect(calls).toHaveLength(1);
    expect(failure).toBeInstanceOf(GraphApiError);
    expect(failure.status).toBe(429);
    expect(failure.message).toMatch(/left alone for 120 seconds/);
  });

  it("takes the ceiling from the caller, so the apply path can wait longer than a read", async () => {
    const tooLong = vi.fn((requestedMs: number) => new Error(`waited out ${requestedMs}`));
    stubFetch(() => jsonResponse({}, 429, { "Retry-After": "5" }));

    await expect(newApi({ retryAfter: { maxWaitMs: 1_000, tooLong } })
      .listItems("site-1", "list-1", { columns: SCHEMA }))
      .rejects.toThrow("waited out 5000");
    expect(tooLong).toHaveBeenCalledWith(5_000);
  });
});
