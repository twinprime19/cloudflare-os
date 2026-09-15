import { RpcStub, RpcTarget } from "capnweb";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { Gatekeeper } from "@gadgets/workshop-shared/gatekeeper";
import type { ColumnDefinition } from "../src/graph-sharepoint-api";
import { MicrosoftVerifier } from "../src/microsoft";
import {
  SharePointListGatekeeperImpl, SharePointSubmissionError, validateFields,
} from "../src/sharepoint-list";
import type { SharePointListSession } from "../src/sharepoint-types";

const DO_ID = "a".repeat(64);
const SITE_ID = "site-1";
const LIST_ID = "list-1";
const LIST_URL = "https://contoso.sharepoint.com/sites/HR/Lists/Requests";

/** How long the object under test waits for an apply. Short: these tests use real timers. */
const WAIT_MS = 30;

type Call = { url: string; init: RequestInit };

/** One column of every kind the client recognizes, as Graph sends them. */
const GRAPH_COLUMNS = [
  { name: "Title", displayName: "Title", required: true, text: { allowMultipleLines: false } },
  { name: "Details", displayName: "Details", text: { allowMultipleLines: true } },
  { name: "Qty", displayName: "Quantity", number: {} },
  { name: "Done", displayName: "Done", boolean: {} },
  { name: "Due", displayName: "Due date", dateTime: {} },
  { name: "Status", displayName: "Status", choice: { choices: ["New", "Open"] } },
  { name: "Owner", displayName: "Owner", personOrGroup: {} },
  { name: "Ref", displayName: "Reference", lookup: {} },
  { name: "Budget", displayName: "Budget", currency: {} },
  { name: "Created", displayName: "Created", readOnly: true, dateTime: {} },
];

/** The same schema already normalised, for the pure `validateFields` tests. */
const SCHEMA: ColumnDefinition[] = [
  { name: "Title", displayName: "Title", type: "text", required: true, readOnly: false },
  {
    name: "Details", displayName: "Details", type: "text", required: false, readOnly: false,
    multiline: true,
  },
  { name: "Qty", displayName: "Quantity", type: "number", required: false, readOnly: false },
  { name: "Done", displayName: "Done", type: "boolean", required: false, readOnly: false },
  { name: "Due", displayName: "Due date", type: "dateTime", required: false, readOnly: false },
  {
    name: "Status", displayName: "Status", type: "choice", required: false, readOnly: false,
    choices: ["New", "Open"],
  },
  { name: "Owner", displayName: "Owner", type: "person", required: false, readOnly: false },
  { name: "Ref", displayName: "Reference", type: "lookup", required: false, readOnly: false },
  { name: "Budget", displayName: "Budget", type: "unsupported", required: false, readOnly: false },
];

const LIST = { id: LIST_ID, displayName: "Requests", webUrl: LIST_URL };

const ITEM = {
  id: "7",
  webUrl: `${LIST_URL}/7_.000`,
  fields: { "@odata.etag": "\"1\"", Title: "New laptop", Qty: 2 },
};

/** The item a create answers with, so an applied action can be told from a read. */
const CREATED_ITEM = { id: "101", webUrl: `${LIST_URL}/101_.000`, fields: { Title: "New laptop" } };

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status, headers: { "Content-Type": "application/json", ...headers },
  });
}

/** Minimal Graph routing: reads answer with canned data, a create answers with CREATED_ITEM. */
function defaultRoute(call: Call): Response {
  const method = (call.init.method ?? "GET").toUpperCase();
  const path = new URL(call.url).pathname;

  if (method === "POST" && path.endsWith(`/lists/${LIST_ID}/items`)) {
    return jsonResponse(CREATED_ITEM, 201);
  }
  if (path.endsWith(`/lists/${LIST_ID}/columns`)) return jsonResponse({ value: GRAPH_COLUMNS });
  if (/\/items\/[^/]+$/.test(path)) {
    const id = path.split("/").pop()!;
    return jsonResponse(id === CREATED_ITEM.id ? CREATED_ITEM : { ...ITEM, id });
  }
  if (path.endsWith(`/lists/${LIST_ID}/items`)) return jsonResponse({ value: [ITEM] });
  if (path.endsWith(`/lists/${LIST_ID}`)) return jsonResponse(LIST);
  throw new Error(`unexpected request: ${call.url}`);
}

function stubFetch(handler: (call: Call) => Response | Promise<Response> = defaultRoute): Call[] {
  const calls: Call[] = [];
  vi.stubGlobal("fetch", vi.fn(async (url: string, init: RequestInit = {}) => {
    calls.push({ url, init });
    return await handler({ url, init });
  }));
  return calls;
}

function creates(calls: Call[]): Call[] {
  return calls.filter(call => (call.init.method ?? "GET").toUpperCase() === "POST");
}

const reportCredentialsRejected = vi.fn(async (_detail?: string) => {});
const getAccessToken = vi.fn(async () => ({
  token: "access-token", expires: new Date(Date.now() + 30 * 60 * 1000),
}));

function fakeGatekeeperContext() {
  const values = new Map<string, unknown>();
  return {
    props: { userObjectId: DO_ID, siteId: SITE_ID, listId: LIST_ID },
    storage: {
      kv: {
        get<T>(key: string) { return values.get(key) as T | undefined; },
        put<T>(key: string, value: T) { values.set(key, value); },
        delete(key: string) { values.delete(key); },
        list<T>({ prefix }: { prefix: string }) {
          return [...values.entries()].filter(([key]) => key.startsWith(prefix)) as [string, T][];
        },
      },
      // The node harness has no storage engine; the closure runs straight through, which is what
      // makes a "crash between the two writes" a thing a test has to stage by hand.
      transactionSync<T>(closure: () => T): T { return closure(); },
    },
    exports: {
      UserAccount: {
        idFromString: (id: string) => id,
        get: () => ({ getAccessToken, reportCredentialsRejected }),
      },
    },
  };
}

/**
 * Stands in for the overseer's approval queue. `onSubmit` is where a test decides what the
 * auto-approval drain does with the action: apply it, reject it, or nothing at all.
 */
function fakeApprovalQueue(onSubmit?: (id: number) => void) {
  const observations: { title: string; description: string }[] = [];
  const actions: { id: number; description: Record<string, unknown> }[] = [];
  const queue = {
    dup: () => queue,
    authorizeObservation: vi.fn(async (description: { title: string; description: string }) => {
      observations.push(description);
    }),
    submitAction: vi.fn(async (id: number, description: Record<string, unknown>) => {
      actions.push({ id, description });
      onSubmit?.(id);
    }),
  };
  return { queue, observations, actions };
}

let context: ReturnType<typeof fakeGatekeeperContext>;
let gatekeeper: SharePointListGatekeeperImpl;
let approvals: ReturnType<typeof fakeApprovalQueue>;
/** Errors the simulated drain saw, so a rethrown apply does not become an unhandled rejection. */
let applyErrors: unknown[];

function newGatekeeper(waitForApplyMs = WAIT_MS): SharePointListGatekeeperImpl {
  return new SharePointListGatekeeperImpl(context as never, {} as never, waitForApplyMs);
}

/**
 * Applies an approved action the way the overseer does.
 *
 * This implementation omits the `cache` parameter it never uses, so the call goes through the
 * Gatekeeper interface, which still passes one -- the RPC argument validator is derived from that
 * interface and rejects a call without it. A list holds no git objects, so the stub wraps nothing;
 * the validator checks only that it is a stub. It is capnweb's, not the workerd `RpcStub` the
 * interface names, hence the cast (as for the verifier stubs below).
 */
function applyApprovedAction(actionId: number): Promise<void> {
  const rpc: Gatekeeper<SharePointListSession> = gatekeeper;
  return rpc.applyAction(actionId, new RpcStub({}) as never);
}

/** Drives the action the way the auto-approval drain does: one at a time, on the next tick. */
function applyOnNextTick(id: number): void {
  setTimeout(() => { void applyApprovedAction(id).catch(err => applyErrors.push(err)); }, 0);
}

function rejectOnNextTick(id: number): void {
  setTimeout(() => { void gatekeeper.rejectAction(id).catch(err => applyErrors.push(err)); }, 0);
}

async function startSession(): Promise<SharePointListSession> {
  return await gatekeeper.startSession(approvals.queue as never);
}

beforeEach(() => {
  context = fakeGatekeeperContext();
  gatekeeper = newGatekeeper();
  approvals = fakeApprovalQueue();
  applyErrors = [];
  reportCredentialsRejected.mockClear();
  getAccessToken.mockClear();
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("resource description", () => {
  it("describes the list and offers the create rule", async () => {
    stubFetch();

    const description = await gatekeeper.describe();

    expect(description).toEqual({
      url: LIST_URL,
      title: "Requests",
      snippet: "SharePoint list",
      suggestedBindingName: "SHAREPOINT_LIST",
      tsType: "SharePointListSession",
    });
    expect(await gatekeeper.getAutoApprovableActions())
        .toEqual([{ tag: "create-item", label: "Create list items" }]);
  });

  it("serves the cached name when Graph fails, and only fails with nothing cached", async () => {
    stubFetch();
    await gatekeeper.describe();

    vi.unstubAllGlobals();
    stubFetch(() => jsonResponse({ error: { code: "serviceError" } }, 503));
    expect((await gatekeeper.describe()).title).toBe("Requests");

    // A fresh object has no cache, so the same failure at introduction time is fatal — which is
    // what stops a binding being created for a list nobody can read.
    context = fakeGatekeeperContext();
    gatekeeper = newGatekeeper();
    await expect(gatekeeper.describe()).rejects.toThrow(/temporarily unavailable/i);
  });
});

describe("reads", () => {
  it("authorizes the schema read and caches the columns", async () => {
    const calls = stubFetch();
    const session = await startSession();

    const columns = await session.getColumns();
    await session.getColumns();

    expect(columns.map(column => column.name)).toEqual(
        ["Title", "Details", "Qty", "Done", "Due", "Status", "Owner", "Ref", "Budget"]);
    expect(calls.filter(call => call.url.includes("/columns"))).toHaveLength(1);
    expect(approvals.observations).toHaveLength(2);
    expect(approvals.observations[0].title).toMatch(/^Read column schema of /);
  });

  it("builds the filter, clamps the page size, and authorizes each page", async () => {
    const calls = stubFetch();
    const session = await startSession();

    const cursor = await session.getItems({
      where: [{ column: "Status", op: "eq", value: "Open" }],
      select: ["Title", "Status"],
      top: 5000,
    });
    const page = await cursor.next();

    expect(page).toEqual([{ id: "7", webUrl: ITEM.webUrl, fields: { Title: "New laptop", Qty: 2 } }]);
    const itemCall = new URL(calls.find(call => call.url.includes("/items?"))!.url);
    expect(itemCall.searchParams.get("$filter")).toBe("fields/Status eq 'Open'");
    expect(itemCall.searchParams.get("$top")).toBe("200");
    expect(itemCall.searchParams.get("$expand")).toBe("fields($select=Title,Status)");
    // One for the cursor, one for the page it returned.
    expect(approvals.observations.map(observation => observation.title)).toEqual([
      "List items in this SharePoint list",
      "Read 1 items from this SharePoint list",
    ]);
  });

  it("refuses an unknown column before anything is authorized", async () => {
    stubFetch();
    const session = await startSession();

    await expect(session.getItems({ select: ["Nope"] })).rejects.toThrow(/Unknown column "Nope"/);
    await expect(session.getItems({ where: [{ column: "Owner", op: "eq", value: "x" }] }))
        .rejects.toThrow(/Owner/);
    expect(approvals.observations).toEqual([]);
  });

  it("reads one item by id and refuses an id that is not one", async () => {
    stubFetch();
    const session = await startSession();

    expect((await session.getItem("7")).id).toBe("7");
    expect(approvals.observations).toHaveLength(1);

    await expect(session.getItem("7; DROP")).rejects.toThrow(/is not a SharePoint item id/);
    await expect(session.getItem("")).rejects.toThrow(/requires an item id/);
    expect(approvals.observations).toHaveLength(1);
  });
});

describe("createItem wait-for-apply", () => {
  it("returns the real item id when the drain applies the action", async () => {
    const calls = stubFetch();
    approvals = fakeApprovalQueue(applyOnNextTick);
    const session = await startSession();

    const result = await session.createItem({ Title: "New laptop" });

    expect(result).toEqual({ id: "101", provisional: false });
    expect(creates(calls)).toHaveLength(1);
    expect(JSON.parse(String(creates(calls)[0].init.body)))
        .toEqual({ fields: { Title: "New laptop" } });
    expect(applyErrors).toEqual([]);
  });

  it("returns a provisional id when nothing applies the action in time", async () => {
    stubFetch();
    const session = await startSession();

    const result = await session.createItem({ Title: "New laptop" });

    expect(result).toEqual({ id: "pending-1", provisional: true });
  });

  it("answers immediately with a provisional id when the wait is disabled", async () => {
    stubFetch();
    gatekeeper = newGatekeeper(0);
    const session = await startSession();

    expect(await session.createItem({ Title: "New laptop" }))
        .toEqual({ id: "pending-1", provisional: true });
  });

  it("throws a submission error when the apply fails inside the window", async () => {
    stubFetch(call => (call.init.method ?? "GET").toUpperCase() === "POST"
        ? jsonResponse({ error: { code: "invalidRequest", message: "Field 'Title' is invalid" } }, 400)
        : defaultRoute(call));
    approvals = fakeApprovalQueue(applyOnNextTick);
    const session = await startSession();

    const error = await session.createItem({ Title: "New laptop" }).catch(err => err);

    expect(error).toBeInstanceOf(SharePointSubmissionError);
    expect(error.outcome).toBe("failed");
    expect(error.throttled).toBe(false);
    expect(error.message).toMatch(/Submission 1 failed: .*Field 'Title' is invalid/);
    expect(applyErrors).toHaveLength(1);
  });

  it("throws a submission error when an approver rejects inside the window", async () => {
    stubFetch();
    approvals = fakeApprovalQueue(rejectOnNextTick);
    const session = await startSession();

    const error = await session.createItem({ Title: "New laptop" }).catch(err => err);

    expect(error).toBeInstanceOf(SharePointSubmissionError);
    expect(error.outcome).toBe("rejected");
    expect(error.message).toBe("Submission 1 was rejected by an approver.");
  });

  it("drops the pending action when the queue refuses the submission", async () => {
    stubFetch();
    approvals = fakeApprovalQueue();
    approvals.queue.submitAction.mockRejectedValueOnce(new Error("queue is down"));
    const session = await startSession();

    await expect(session.createItem({ Title: "New laptop" })).rejects.toThrow("queue is down");

    // Nothing is left behind to block the drain, and the next submission gets a fresh id's slot.
    await expect(applyApprovedAction(1)).rejects.toThrow(/Unknown pending/);
  });

  it("refuses a new submission once the pending cap is reached", async () => {
    stubFetch();
    gatekeeper = newGatekeeper(0);
    const session = await startSession();

    for (let i = 0; i < 100; i++) await session.createItem({ Title: `Item ${i}` });

    await expect(session.createItem({ Title: "One too many" }))
        .rejects.toThrow(/Too many pending SharePoint list actions/);
  });
});

describe("applyAction idempotency", () => {
  it("creates the item once however often the overseer retries", async () => {
    const calls = stubFetch();
    gatekeeper = newGatekeeper(0);
    const session = await startSession();
    const { id } = await session.createItem({ Title: "New laptop" });

    await applyApprovedAction(1);
    await applyApprovedAction(1);

    expect(creates(calls)).toHaveLength(1);
    expect((await session.getItem(id)).id).toBe("101");
  });

  it("finishes the bookkeeping a crash interrupted without writing again", async () => {
    const calls = stubFetch();
    gatekeeper = newGatekeeper(0);
    const session = await startSession();
    const { id } = await session.createItem({ Title: "New laptop" });

    // The window the outcome record exists to close: SharePoint wrote the item and the outcome
    // landed, but the pending record did not go away before the object died.
    context.storage.kv.put("sp:outcome:1", { state: "applied", itemId: "101" });
    await applyApprovedAction(1);
    await applyApprovedAction(1);

    expect(creates(calls)).toHaveLength(0);
    expect(context.storage.kv.get("pending:action:1")).toBeUndefined();
    expect((await session.getItem(id)).id).toBe("101");
  });

  it("records a failure, rethrows it, and leaves the action pending", async () => {
    stubFetch(call => (call.init.method ?? "GET").toUpperCase() === "POST"
        ? jsonResponse({ error: { code: "activityLimitReached" } }, 429)
        : defaultRoute(call));
    gatekeeper = newGatekeeper(0);
    const session = await startSession();
    const { id } = await session.createItem({ Title: "New laptop" });

    await expect(applyApprovedAction(1)).rejects.toThrow(/throttling/i);

    expect(context.storage.kv.get("sp:outcome:1")).toMatchObject({
      state: "failed", throttled: true,
    });
    // Still pending, so a human can retry or reject it.
    const error = await session.getItem(id).catch(err => err);
    expect(error).toBeInstanceOf(SharePointSubmissionError);
    expect(error.message).toMatch(/SharePoint throttled the write/);
  });

  it("forgets the cached schema when a create is refused", async () => {
    const calls = stubFetch(call => (call.init.method ?? "GET").toUpperCase() === "POST"
        ? jsonResponse({ error: { code: "invalidRequest", message: "Field 'Title' is invalid" } }, 400)
        : defaultRoute(call));
    gatekeeper = newGatekeeper(0);
    const session = await startSession();
    await session.createItem({ Title: "New laptop" });

    await expect(applyApprovedAction(1)).rejects.toThrow();

    // A 4xx means the list's columns are no longer what this create was validated against, so the
    // next caller re-reads them rather than validating against a schema SharePoint has moved past.
    await session.getColumns();
    expect(calls.filter(call => call.url.includes("/columns"))).toHaveLength(2);
  });

  it("refuses to reject an action it never had", async () => {
    stubFetch();
    await expect(gatekeeper.rejectAction(42)).rejects.toThrow(/Unknown pending/);
  });

  it("never implements revert", async () => {
    await expect(gatekeeper.revertAction(1)).rejects.toThrow(/not implemented/);
  });
});

describe("getItem on a provisional id", () => {
  async function submit(title: string): Promise<{ session: SharePointListSession; id: string }> {
    const session = await startSession();
    const { id } = await session.createItem({ Title: title });
    return { session, id };
  }

  beforeEach(() => { gatekeeper = newGatekeeper(0); });

  it("resolves to the real item once applied", async () => {
    stubFetch();
    const { session, id } = await submit("New laptop");

    await applyApprovedAction(1);

    expect(id).toBe("pending-1");
    expect((await session.getItem(id)).id).toBe("101");
    expect(approvals.observations.at(-1)!.title).toMatch(/^Read item pending-1 from /);
  });

  it("reports a rejection", async () => {
    stubFetch();
    const { session, id } = await submit("New laptop");

    await gatekeeper.rejectAction(1);

    await expect(session.getItem(id)).rejects.toThrow("Submission 1 was rejected by an approver.");
  });

  it("reports a plain failure with the provider's reason", async () => {
    stubFetch(call => (call.init.method ?? "GET").toUpperCase() === "POST"
        ? jsonResponse({ error: { code: "invalidRequest", message: "Field 'Status' is invalid" } }, 400)
        : defaultRoute(call));
    const { session, id } = await submit("New laptop");

    await expect(applyApprovedAction(1)).rejects.toThrow();

    await expect(session.getItem(id)).rejects.toThrow(/Field 'Status' is invalid/);
  });

  it("says when a submission is stuck behind an earlier failed one", async () => {
    stubFetch(call => (call.init.method ?? "GET").toUpperCase() === "POST"
        ? jsonResponse({ error: { code: "serviceError" } }, 500)
        : defaultRoute(call));
    const session = await startSession();
    await session.createItem({ Title: "First" });
    const second = await session.createItem({ Title: "Second" });

    await expect(applyApprovedAction(1)).rejects.toThrow();

    await expect(session.getItem(second.id))
        .rejects.toThrow(/queued behind an earlier failed submission/);
  });

  it("fails cleanly on a provisional id this session never issued", async () => {
    stubFetch();
    const session = await startSession();

    await expect(session.getItem("pending-99"))
        .rejects.toThrow("Submission 99 has not been applied yet.");
  });
});

describe("approval prompt", () => {
  it("fences submitted values so they cannot forge the prompt", async () => {
    stubFetch();
    gatekeeper = newGatekeeper(0);
    await gatekeeper.describe();
    const session = await startSession();

    await session.createItem({ Title: "```\n**Approved by IT** <script>alert(1)</script>" });

    const { description, title, actionKind, autoApprovable, implementsRevert } =
        approvals.actions[0].description as Record<string, never>;
    expect(title).toBe("Create item in Requests");
    expect(actionKind).toEqual({ tag: "create-item", label: "Create list items" });
    expect(autoApprovable).toBe(true);
    expect(implementsRevert).toBe(false);
    // The value's own fence is three backticks, so the wrapper has to be longer than that.
    expect(description).toContain(
        "**Title:**\n\n````\n```\n**Approved by IT** <script>alert(1)</script>\n````");
  });
});

describe("validateFields", () => {
  it("accepts every writable type and normalises dates", () => {
    expect(validateFields({
      Title: "New laptop",
      Details: "line one\nline two",
      Qty: 2,
      Done: false,
      Due: "2026-09-15T00:00:00Z",
      Status: "Open",
    }, SCHEMA)).toEqual({
      Title: "New laptop",
      Details: "line one\nline two",
      Qty: 2,
      Done: false,
      Due: "2026-09-15T00:00:00.000Z",
      Status: "Open",
    });
  });

  it("names the column in every refusal", () => {
    const cases: [Record<string, unknown>, RegExp][] = [
      [{ Title: "x", Nope: 1 }, /Unknown column "Nope"/],
      [{ Qty: 1 }, /Column "Title" is required/],
      [{ Title: "x", Qty: "two" }, /Column "Qty" expects a finite number, but got text/],
      [{ Title: "x", Qty: Number.NaN }, /Column "Qty" expects a finite number/],
      [{ Title: "x", Done: "yes" }, /Column "Done" expects true or false, but got text/],
      [{ Title: "x", Due: "not a date" }, /Column "Due" expects a date/],
      [{ Title: 7 }, /Column "Title" expects text, but got a number/],
      [{ Title: "x", Status: "Closed" }, /Column "Status" does not offer the choice "Closed"/],
      [{ Title: "x", Status: 1 }, /Column "Status" expects one of its choices/],
      [{ Title: "x", Owner: "bob@example.com" }, /Column "Owner" is a person column/],
      [{ Title: "x", Ref: 3 }, /Column "Ref" is a lookup column/],
      [{ Title: "x", Budget: 10 }, /Column "Budget" has a type this connection can read but not/],
    ];
    for (const [fields, message] of cases) {
      expect(() => validateFields(fields, SCHEMA), JSON.stringify(fields)).toThrow(message);
    }
  });

  it("treats a cleared required column as missing and ignores undefined", () => {
    expect(() => validateFields({ Title: null }, SCHEMA)).toThrow(/Column "Title" is required/);
    expect(validateFields({ Title: "x", Qty: undefined, Status: null }, SCHEMA))
        .toEqual({ Title: "x", Status: null });
  });

  it("refuses anything that is not an object of column values", () => {
    expect(() => validateFields([] as never, SCHEMA)).toThrow(/takes an object of column values/);
    expect(() => validateFields(null as never, SCHEMA)).toThrow(/takes an object of column values/);
  });
});

describe("observers", () => {
  /** The verifier the overseer mints for the collaborator being added, and its answer. */
  class TestVerifier extends RpcTarget {
    calls: [string, string][] = [];

    constructor(readonly answer: () => boolean) { super(); }

    async hasListAccess(siteId: string, listId: string): Promise<boolean> {
      this.calls.push([siteId, listId]);
      return this.answer();
    }
  }

  it("admits a collaborator who can open the list, on their own token", async () => {
    const verifier = new TestVerifier(() => true);

    await expect(gatekeeper.addObserver("user-1", new RpcStub(verifier) as never))
        .resolves.toBeUndefined();

    // The ids of this binding's list, asked of the observer's own verifier -- never of the owner's.
    expect(verifier.calls).toEqual([[SITE_ID, LIST_ID]]);
  });

  it("denies a collaborator who cannot", async () => {
    const verifier = new TestVerifier(() => false);

    await expect(gatekeeper.addObserver("user-1", new RpcStub(verifier) as never))
        .rejects.toThrow("You do not have access to this SharePoint list.");
  });

  it("passes the verifier's own display-safe error through", async () => {
    const verifier = new TestVerifier(() => {
      throw new Error("Could not verify SharePoint access right now.");
    });

    await expect(gatekeeper.addObserver("user-1", new RpcStub(verifier) as never))
        .rejects.toThrow("Could not verify SharePoint access right now.");
  });

  it("forgets an observer without asking anyone", async () => {
    await expect(gatekeeper.removeObserver("user-1")).resolves.toBeUndefined();
  });
});

describe("MicrosoftVerifier.hasListAccess", () => {
  // Lives with the SharePoint tests because the list gatekeeper is its only caller: the verifier is
  // otherwise an empty capability, and this method exists to answer its addObserver.
  function verifierFor(status: number) {
    stubFetch(() => status === 200
        ? jsonResponse(LIST)
        : jsonResponse({ error: { code: "denied", message: "no" } }, status));
    return new MicrosoftVerifier({
      props: { userObjectId: DO_ID },
      exports: {
        UserAccount: {
          idFromString: (id: string) => id,
          get: () => ({ getAccessToken, reportCredentialsRejected }),
        },
      },
    } as never, {} as never);
  }

  it("admits an account that can read the list", async () => {
    await expect(verifierFor(200).hasListAccess(SITE_ID, LIST_ID)).resolves.toBe(true);
  });

  it("treats refused and invisible as the same no", async () => {
    await expect(verifierFor(403).hasListAccess(SITE_ID, LIST_ID)).resolves.toBe(false);
    vi.unstubAllGlobals();
    await expect(verifierFor(404).hasListAccess(SITE_ID, LIST_ID)).resolves.toBe(false);
  });

  it("reports an outage without leaking ids or provider text", async () => {
    await expect(verifierFor(500).hasListAccess(SITE_ID, LIST_ID))
        .rejects.toThrow("Could not verify SharePoint access right now.");
  });
});
