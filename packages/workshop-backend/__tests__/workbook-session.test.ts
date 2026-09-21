// The programmatic Workbook binding (workbook-session.ts): what the agent reaches when it reads a
// spreadsheet attachment's rows from executeCode.
//
// What is pinned here: the rows come back exactly as the spreadsheet held them, including across
// the boundary between two stored chunks (the one place a paging bug would show as silently
// missing or repeated rows); ranges are validated with errors that tell the agent how to page; the
// session opens only for the chat that owns the attachment, and never for a gadget; the binding's
// description carries the same API text the .d.ts declares; and the binding name derived from a
// file name is deterministic.
//
// Runs the real OverseerImpl inside workerd over real Durable Object storage, the harness
// chat-attachment-import.test.ts uses, with the attachment put in place through the real import
// pipeline so the records under test are the ones production writes.

import { describe, expect, it } from "vitest";
import { env } from "cloudflare:workers";
import { runInDurableObject } from "cloudflare:test";
import * as XLSX from "xlsx";
import type { AiChatMessage, ChatAttachmentUpload } from "@gadgets/workshop-shared/api";
import type { OverseerDurableObject } from "../src/overseer.js";
import { MAX_WORKBOOK_ROWS_PER_CALL } from "../src/workbook-session";
import { workbookBindingName } from "../src/chat-attachment-workbook";
// Vite's ?raw import; resolved relative to this file, so the test below can compare the shipped
// text module against its source of truth.
// @ts-expect-error -- ?raw imports have no type declaration
import WORKBOOK_DTS_SOURCE from "../src/workbook-binding.d.ts?raw";
// The text module workshop-backend ships -- a symlink to the .d.ts above (see
// describeAttachmentBinding in overseer.ts).
import WORKBOOK_BINDING_TYPES from "../src/workbook-binding.txt";

declare module "cloudflare:workers" {
  interface ProvidedEnv {
    TEST_OVERSEER: DurableObjectNamespace<OverseerDurableObject>;
  }
}

const CHAT_ID = 1;
const OTHER_CHAT_ID = 2;
const OWNER_PROFILE_ID = "alice@example.com";
const SOURCE_LABEL = "Outlook Mailbox";
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** The width of a Ledger row's text column: enough that 1,200 rows need more than one chunk. */
const LEDGER_FILLER = "x".repeat(900);
const LEDGER_ROWS = 1_200;

let doCounter = 0;

/** A workspace with two chats and a known owner, which is all the import needs. */
async function withChat(fn: (impl: any) => Promise<void>): Promise<void> {
  let stub = env.TEST_OVERSEER.getByName(`workbook-session-${++doCounter}`);
  await runInDurableObject(stub, async (instance: OverseerDurableObject) => {
    let impl = (instance as unknown as { impl: any }).impl;
    // The author's id is the workspace owner's profile id; seeding the cache stands in for the
    // owner Durable Object this unit-test worker does not run.
    impl.ownerProfileId = OWNER_PROFILE_ID;
    for (let id of [CHAT_ID, OTHER_CHAT_ID]) {
      impl.storage.chatMeta.put(
          { id, title: "Chat", started: new Date(0), lastActive: new Date(id) });
    }
    await fn(impl);
  });
}

function buildWorkbook(sheets: { name: string, rows: unknown[][] }[]): Uint8Array {
  let workbook = XLSX.utils.book_new();
  for (let sheet of sheets) {
    XLSX.utils.book_append_sheet(
        workbook, XLSX.utils.aoa_to_sheet(sheet.rows, { cellDates: true }), sheet.name);
  }
  return new Uint8Array(XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as ArrayBuffer);
}

// A small sheet to read whole, and one whose rows span more than one stored chunk. The date is
// built from local components because SheetJS's *writer* reads a Date by its local clock (its
// reader materializes the cell's wall clock in UTC components), so this is the cell a spreadsheet
// would hold on any host.
function workbookUpload(name = "quarterly.xlsx"): ChatAttachmentUpload {
  let summary = [
    ["KPI", "Value"],
    ["Revenue", 431000],
    ["As of", new Date(2026, 7, 6)],
  ];
  let ledger = Array.from({ length: LEDGER_ROWS }, (_, index) => [index, LEDGER_FILLER]);
  return {
    mimeType: XLSX_MIME_TYPE,
    content: buildWorkbook([{ name: "Summary", rows: summary }, { name: "Ledger", rows: ledger }]),
    name,
  };
}

/** A workbook with more rows than one getRows() call may return, for the paging cap. */
function tallWorkbookUpload(): ChatAttachmentUpload {
  let rows = Array.from({ length: MAX_WORKBOOK_ROWS_PER_CALL + 1 }, (_, index) => [index]);
  return {
    mimeType: XLSX_MIME_TYPE,
    content: buildWorkbook([{ name: "Tall", rows }]),
    name: "tall.xlsx",
  };
}

/** The single message an import posts. */
function importedMessage(impl: any): AiChatMessage {
  return [...impl.storage.chats.list()].find((msg: AiChatMessage) => msg.type === "message")!;
}

type WorkbookFixture = {
  impl: any;
  fileId: string;
  /** The stored index the session pages against (WorkbookMeta). */
  meta: any;
  session: any;
};

/** Imports a workbook into CHAT_ID and opens the binding its agent would get. */
async function withWorkbook(
    fn: (fixture: WorkbookFixture) => Promise<void>,
    upload: ChatAttachmentUpload = workbookUpload()): Promise<void> {
  await withChat(async impl => {
    await impl.importChatAttachments(CHAT_ID, [upload], SOURCE_LABEL, "anthropic");
    let fileId = importedMessage(impl).attachments![0].id;
    let meta = impl.storage.chatAttachmentContent.get(fileId).workbook;
    let session = await impl.startGatekeeperSession(
        { type: "workbook", id: fileId }, { from: "agent", chatId: CHAT_ID });
    await fn({ impl, fileId, meta, session });
  });
}

describe("Workbook binding", () => {
  it("lists every sheet with the dimensions the parse measured", () => withWorkbook(
      async ({ session }) => {
    expect(await session.listSheets()).toEqual([
      { name: "Summary", rowCount: 3, colCount: 2 },
      { name: "Ledger", rowCount: LEDGER_ROWS, colCount: 2 },
    ]);
  }));

  it("returns a whole sheet from row 0, dates as ISO strings", () => withWorkbook(
      async ({ session }) => {
    // Row 0 is the spreadsheet's first row -- the header, here -- not the first row *under* a
    // header: nothing in this path treats any row as special.
    expect(await session.getRows("Summary")).toEqual([
      ["KPI", "Value"],
      ["Revenue", 431000],
      ["As of", "2026-08-06"],
    ]);
  }));

  it("reads a range spanning two stored chunks as one contiguous run", () => withWorkbook(
      async ({ session, meta }) => {
    // The Ledger sheet is chunked, and this is the seam: a chunk-selection or slicing bug shows
    // up here as a missing or repeated row and nowhere else.
    let boundary = meta.sheets[1].chunkRowStarts[1];
    expect(boundary).toBeGreaterThan(0);
    expect(meta.sheets[1].chunkRowStarts.length).toBeGreaterThan(1);

    let rows = await session.getRows("Ledger", boundary - 2, boundary + 2);
    expect(rows.map((row: unknown[]) => row[0]))
        .toEqual([boundary - 2, boundary - 1, boundary, boundary + 1]);
    expect(rows.every((row: unknown[]) => row[1] === LEDGER_FILLER)).toBe(true);
  }));

  it("pages a chunked sheet without gaps or repeats", () => withWorkbook(
      async ({ session, meta }) => {
    let boundary = meta.sheets[1].chunkRowStarts[1];
    let whole = await session.getRows("Ledger");
    expect(whole.length).toBe(LEDGER_ROWS);
    expect(whole.map((row: unknown[]) => row[0]))
        .toEqual(Array.from({ length: LEDGER_ROWS }, (_, index) => index));

    // Pages that split inside a chunk and pages that split on the chunk seam both rebuild it.
    let head = await session.getRows("Ledger", 0, 500);
    let tail = await session.getRows("Ledger", 500);
    expect([...head, ...tail]).toEqual(whole);
    expect([...await session.getRows("Ledger", 0, boundary),
            ...await session.getRows("Ledger", boundary)]).toEqual(whole);

    // A start at or past the end is what ends a paging loop, not an error.
    expect(await session.getRows("Ledger", LEDGER_ROWS)).toEqual([]);
    expect(await session.getRows("Ledger", LEDGER_ROWS + 10)).toEqual([]);
    // An end past the sheet is clamped to it, so the last page need not know where it stops.
    expect(await session.getRows("Ledger", LEDGER_ROWS - 2, LEDGER_ROWS + 1000))
        .toEqual(whole.slice(LEDGER_ROWS - 2));
  }));

  it("refuses a sheet the workbook does not have, naming the ones it does", () => withWorkbook(
      async ({ session }) => {
    await expect(session.getRows("Sales")).rejects
        .toThrow(`This workbook has no sheet named "Sales". Its sheets are: "Summary", "Ledger".`);
  }));

  it("refuses a range that is not whole rows, or runs backwards", () => withWorkbook(
      async ({ session }) => {
    await expect(session.getRows("Summary", 0, 1.5)).rejects.toThrow(/whole numbers of rows/);
    await expect(session.getRows("Summary", Number.NaN)).rejects.toThrow(/whole numbers of rows/);
    await expect(session.getRows("Summary", -1)).rejects.toThrow(/must be 0 or greater/);
    await expect(session.getRows("Summary", 2, 1)).rejects
        .toThrow(/end \(1\) is before start \(2\)/);
  }));

  it("refuses more rows than one call may return, and takes the same range in pages", () =>
      withWorkbook(async ({ session }) => {
    await expect(session.getRows("Tall", 0, MAX_WORKBOOK_ROWS_PER_CALL + 1)).rejects
        .toThrow(`That range is ${MAX_WORKBOOK_ROWS_PER_CALL + 1} rows, and at most ` +
            `${MAX_WORKBOOK_ROWS_PER_CALL} can be read per call.`);
    // The whole sheet defaults to the same over-cap range, so it is refused the same way.
    await expect(session.getRows("Tall")).rejects.toThrow(/Read the sheet in pages/);

    let page = await session.getRows("Tall", 0, MAX_WORKBOOK_ROWS_PER_CALL);
    expect(page.length).toBe(MAX_WORKBOOK_ROWS_PER_CALL);
    expect(page[MAX_WORKBOOK_ROWS_PER_CALL - 1]).toEqual([MAX_WORKBOOK_ROWS_PER_CALL - 1]);
    expect(await session.getRows("Tall", MAX_WORKBOOK_ROWS_PER_CALL))
        .toEqual([[MAX_WORKBOOK_ROWS_PER_CALL]]);
  }, tallWorkbookUpload()));

  it("opens only for the agent of the chat that owns the attachment", () => withWorkbook(
      async ({ impl, fileId }) => {
    // The refusals are thrown where the session is dispatched, as the worktree binding's are.
    let open = (caller: unknown) =>
        impl.startGatekeeperSession({ type: "workbook", id: fileId }, caller);

    // A gadget must never page through a user's spreadsheet, even holding a stub the agent passed
    // it; and an attachment belongs to the history of exactly one chat.
    expect(() => open({ from: "gadget", gadgetId: 1, chatId: CHAT_ID }))
        .toThrow("Workbook bindings are only available to the agent's executeCode.");
    expect(() => open({ from: "user", chatId: CHAT_ID }))
        .toThrow("Workbook bindings are only available to the agent's executeCode.");
    expect(() => open({ from: "agent", chatId: OTHER_CHAT_ID }))
        .toThrow("This workbook is no longer attached to this chat.");

    // And nothing opens once the attachment is gone.
    impl.deleteChatAttachmentContent(fileId);
    expect(() => open({ from: "agent", chatId: CHAT_ID }))
        .toThrow("This workbook is no longer attached to this chat.");
  }));

  it("keeps serving a session minted by an earlier execution", () => withWorkbook(
      async ({ session }) => {
    // Deliberately unlike the worktree binding: an attachment's rows are written once and never
    // edited, so there is no turn state a retained stub could revive against -- only the same
    // immutable rows, which the agent may read at any point in its turn.
    expect((await session.getRows("Summary")).length).toBe(3);
    expect((await session.getRows("Summary")).length).toBe(3);
  }));

  it("builds the agent's env from an attachment binding, for the owning chat only", () =>
      withWorkbook(async ({ impl, fileId }) => {
    let bindings = { QUARTERLY_XLSX: { type: "attachment", id: fileId } };
    expect(impl.getEnvForAgent(CHAT_ID, bindings, "execution-1").QUARTERLY_XLSX).toBeDefined();
    // Another chat's entry, and a deleted attachment, are skipped the way a deleted gadget is:
    // the run proceeds without the name rather than failing.
    expect(impl.getEnvForAgent(OTHER_CHAT_ID, bindings, "execution-1").QUARTERLY_XLSX)
        .toBeUndefined();
    impl.deleteChatAttachmentContent(fileId);
    expect(impl.getEnvForAgent(CHAT_ID, bindings, "execution-1").QUARTERLY_XLSX).toBeUndefined();
  }));
});

describe("describeAttachmentBinding", () => {
  it("names the file, the sheets and their shape, and serves the agent API", () => withWorkbook(
      async ({ impl, fileId }) => {
    let description = await impl.describeAttachmentBinding(CHAT_ID, "QUARTERLY_XLSX", fileId);

    expect(description).toContain("Binding: QUARTERLY_XLSX");
    // The name came from outside the workspace, so it is fenced the way every agent-facing
    // untrusted file name is, rather than quoted.
    expect(description).toContain('the workbook `quarterly.xlsx` attached to this chat');
    expect(description).toContain(
        `2 sheets — "Summary" (3 rows × 2 cols), "Ledger" (1,200 rows × 2 cols)`);
    expect(description).toContain("env.QUARTERLY_XLSX provides the following API");
    expect(description).toContain("export interface Workbook");
    expect(description).toContain("getRows(sheet: string, start?: number, end?: number)");
    // Only the agent-facing section ships; the file header (and marker) stay out.
    expect(description).not.toContain("BEGIN AGENT API");
  }));

  it("refuses an attachment that is not this chat's", () => withWorkbook(
      async ({ impl, fileId }) => {
    await expect(impl.describeAttachmentBinding(OTHER_CHAT_ID, "QUARTERLY_XLSX", fileId))
        .rejects.toThrow("The attachment behind QUARTERLY_XLSX is no longer available.");
    impl.deleteChatAttachmentContent(fileId);
    await expect(impl.describeAttachmentBinding(CHAT_ID, "QUARTERLY_XLSX", fileId))
        .rejects.toThrow("The attachment behind QUARTERLY_XLSX is no longer available.");
  }));

  it("workbook-binding.txt resolves to workbook-binding.d.ts", () => {
    // The .txt is a symlink to the .d.ts (same-directory: the validate build only materializes
    // files whose real location is inside the package), so identity is structural; this guards the
    // link itself and the text-module pipeline that ships it.
    expect(WORKBOOK_BINDING_TYPES).toBe(WORKBOOK_DTS_SOURCE);
  });
});

describe("workbookBindingName", () => {
  // Replay rebuilds the chat's binding map from the log on every turn, so this derivation is
  // pinned: a name that moved between turns would leave the conversation's `env.NAME` references
  // pointing at the wrong workbook, or at nothing.
  it("derives a stable identifier from a file name", () => {
    let none: ReadonlySet<string> = new Set();
    expect(workbookBindingName("big.xlsx", none)).toBe("big_xlsx");
    expect(workbookBindingName("Báo cáo Q3.xlsx", none)).toBe("B_o_c_o_Q3_xlsx");
    expect(workbookBindingName("2026.xls", none)).toBe("file_2026_xls");
    expect(workbookBindingName("class.xlsx", none)).toBe("class_xlsx");
    // Bare reserved words and Object.prototype members cannot be binding names on their own.
    expect(workbookBindingName("class", none)).toBe("file_class");
    expect(workbookBindingName("constructor", none)).toBe("file_constructor");
    // Nothing usable in the name at all, including an attachment that arrived without one.
    expect(workbookBindingName(undefined, none)).toBe("file");
    expect(workbookBindingName("- — -", none)).toBe("file");
  });

  it("dedupes against the names already in scope, in order", () => {
    expect(workbookBindingName("big.xlsx", new Set(["big_xlsx"]))).toBe("big_xlsx_2");
    expect(workbookBindingName("big.xlsx", new Set(["big_xlsx", "big_xlsx_2"])))
        .toBe("big_xlsx_3");
    expect(workbookBindingName(undefined, new Set(["file"]))).toBe("file_2");
  });
});
