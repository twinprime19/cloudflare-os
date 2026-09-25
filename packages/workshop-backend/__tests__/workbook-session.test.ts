// The programmatic Workbook binding (workbook-session.ts): what the agent reaches when it reads a
// spreadsheet attachment's rows from executeCode.
//
// What is pinned here: the rows come back exactly as the spreadsheet held them, including across
// the boundary between two stored chunks (the one place a paging bug would show as silently
// missing or repeated rows); ranges are validated with errors that tell the agent how to page; the
// session opens only for the chat that owns the attachment, and never for a gadget; the binding's
// description carries the same API text the .d.ts declares; and readSheet's ranges, caps and
// rendering -- the text the model reads cells from -- come out exactly as specified.
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
import {
  MAX_READ_SHEET_BYTES, MAX_READ_SHEET_ROWS, MAX_WORKBOOK_ROWS_PER_CALL, parseSheetRange,
  readSheetRows,
} from "../src/workbook-session";
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

/**
 * A wide sheet with a blank row between its header and its data, and a numeric run longer than the
 * attachment text would ever spell out.
 */
function wideWorkbookUpload(): ChatAttachmentUpload {
  let years = Array.from({ length: 12 }, (_, index) => 2020 + index);
  let values = Array.from({ length: 12 }, (_, index) => index + 1);
  return {
    mimeType: XLSX_MIME_TYPE,
    content: buildWorkbook([{ name: "Plan", rows: [["Label", ...years], [], ["Revenue", ...values]] }]),
    name: "plan.xlsx",
  };
}

/** A blank first row, then one row whose cells together exceed the readSheet byte cap. */
function oversizedRowWorkbookUpload(): ChatAttachmentUpload {
  let cells = Array.from({ length: 30 }, (_, index) => `${index}-${"y".repeat(1_000)}`);
  return {
    mimeType: XLSX_MIME_TYPE,
    content: buildWorkbook([{ name: "Wide", rows: [[], cells, ["after"]] }]),
    name: "wide.xlsx",
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

describe("parseSheetRange", () => {
  // Shaped like the sheet the tool was designed against: 394 rows, columns A–AV.
  const SHEET = { rowCount: 394, colCount: 48 };

  it.each([
    [undefined, { from: 0, to: 50, colStart: 0, colEnd: 48 }],
    ["", { from: 0, to: 50, colStart: 0, colEnd: 48 }],
    ["150:160", { from: 149, to: 160, colStart: 0, colEnd: 48 }],
    ["A150:AV160", { from: 149, to: 160, colStart: 0, colEnd: 48 }],
    [" c5:e9 ", { from: 4, to: 9, colStart: 2, colEnd: 5 }],
    ["7:7", { from: 6, to: 7, colStart: 0, colEnd: 48 }],
    ["1:200", { from: 0, to: 200, colStart: 0, colEnd: 48 }],
    // Ends past the sheet clamp to it, rows and columns alike -- so does a page that would be
    // over the cap if it were not clamped first.
    ["B390:AZ500", { from: 389, to: 394, colStart: 1, colEnd: 48 }],
    ["300:600", { from: 299, to: 394, colStart: 0, colEnd: 48 }],
  ])("reads %j as %j", (range, expected) => {
    expect(parseSheetRange(range, SHEET)).toEqual(expected);
  });

  it.each(["A1-B2", "A:B", "1", "A1:B", "1:B2", "AAAA1:B2", "1:2:3"])(
      "refuses the malformed range %j, showing both forms", range => {
    expect(() => parseSheetRange(range, SHEET)).toThrow(
        `Invalid range ${JSON.stringify(range)}. Use "A150:AV160" (cells) or "150:160" ` +
        `(whole rows); rows count from 1 and both ends are included.`);
  });

  it("refuses row 0, backwards ranges, and a start past the sheet", () => {
    expect(() => parseSheetRange("0:5", SHEET)).toThrow(`Invalid range "0:5": rows count from 1.`);
    expect(() => parseSheetRange("160:150", SHEET)).toThrow(
        `Invalid range "160:150": it runs backwards. Write the top-left cell first, ` +
        `e.g. "A150:AV160".`);
    expect(() => parseSheetRange("C1:A5", SHEET)).toThrow(/it runs backwards/);
    expect(() => parseSheetRange("400:410", SHEET)).toThrow(
        "Row 400 is past the end of this sheet, which has 394 rows.");
    expect(() => parseSheetRange("AW1:AX2", SHEET)).toThrow(
        "Column AW is past the last column of this sheet, AV.");
  });

  it("refuses more than 200 rows with the page to read instead", () => {
    expect(MAX_READ_SHEET_ROWS).toBe(200);
    expect(() => parseSheetRange("1:201", SHEET)).toThrow(
        `That range is 201 rows; readSheet returns at most 200 per call. Read it in pages, ` +
        `e.g. "1:200".`);
    expect(() => parseSheetRange("A150:AV394", SHEET)).toThrow(
        `That range is 245 rows; readSheet returns at most 200 per call. Read it in pages, ` +
        `e.g. "150:349".`);
  });
});

describe("readWorkbookRange", () => {
  it("renders a header line and every cell with its address", () => withWorkbook(
      async ({ impl, fileId }) => {
    expect(await impl.readWorkbookRange(CHAT_ID, "QUARTERLY_XLSX", fileId, "Summary", undefined))
        .toBe(`Sheet "Summary" rows 1–3 of 3, columns A–B (blank rows omitted)\n` +
            `1 A=KPI B=Value\n` +
            `2 A=Revenue B=431000\n` +
            `3 A="As of" B=2026-08-06`);
  }));

  it("omits blank rows, never collapses a numeric run, and keeps real column letters", () =>
      withWorkbook(async ({ impl, fileId }) => {
    let whole = await impl.readWorkbookRange(CHAT_ID, "PLAN_XLSX", fileId, "Plan", "1:3");
    let lines = whole.split("\n");
    expect(lines[0]).toBe(`Sheet "Plan" rows 1–3 of 3, columns A–M (blank rows omitted)`);
    // Row 2 is blank and so has no line; row 3's twelve numbers are each written out, where the
    // attachment text would have collapsed them to one range entry.
    expect(lines.slice(1)).toEqual([
      "1 A=Label B=2020 C=2021 D=2022 E=2023 F=2024 G=2025 H=2026 I=2027 J=2028 K=2029 L=2030 " +
          "M=2031",
      "3 A=Revenue B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12",
    ]);
    expect(whole).not.toContain("nums");

    // A column span keeps each cell's own column letter rather than renumbering from A.
    expect(await impl.readWorkbookRange(CHAT_ID, "PLAN_XLSX", fileId, "Plan", "C1:E3"))
        .toBe(`Sheet "Plan" rows 1–3 of 3, columns C–E (blank rows omitted)\n` +
            `1 C=2021 D=2022 E=2023\n` +
            `3 C=2 D=3 E=4`);
  }, wideWorkbookUpload()));

  it("stops at 24 KiB on a whole row and names the range to continue with", () => withWorkbook(
      async ({ impl, fileId }) => {
    let text: string = await impl.readWorkbookRange(
        CHAT_ID, "QUARTERLY_XLSX", fileId, "Ledger", "1:200");
    let lines = text.split("\n");
    let notice = lines.at(-1)!;
    let rowLines = lines.slice(1, -1);
    let lastRow = Number(rowLines.at(-1)!.split(" ")[0]);

    expect(lines[0]).toBe(`Sheet "Ledger" rows 1–200 of 1,200, columns A–B (blank rows omitted)`);
    expect(notice).toBe(`… truncated after row ${lastRow} (24 KiB). Continue with range ` +
        `${lastRow + 1}:200.`);
    // Whole rows, in order, from the first, within the budget.
    expect(rowLines.map(line => Number(line.split(" ")[0])))
        .toEqual(Array.from({ length: lastRow }, (_, index) => index + 1));
    expect(rowLines.every(line => line.endsWith(`B=${LEDGER_FILLER}`))).toBe(true);
    expect(new TextEncoder().encode(rowLines.join("\n")).byteLength)
        .toBeLessThanOrEqual(MAX_READ_SHEET_BYTES);
    expect(lastRow).toBeLessThan(200);

    // The continuation picks up exactly where the page stopped.
    let next: string = await impl.readWorkbookRange(
        CHAT_ID, "QUARTERLY_XLSX", fileId, "Ledger", `${lastRow + 1}:200`);
    expect(next.split("\n")[1].startsWith(`${lastRow + 1} A=${lastRow} `)).toBe(true);

    // A column-limited page continues with the same columns.
    let narrow: string = await impl.readWorkbookRange(
        CHAT_ID, "QUARTERLY_XLSX", fileId, "Ledger", "B1:B200");
    expect(narrow.split("\n").at(-1)).toMatch(/Continue with range B\d+:B200\.$/);
  }));

  it("tells the model to read fewer columns when one row alone is over the cap", () =>
      withWorkbook(async ({ impl, fileId }) => {
    // Row 1 is blank and costs nothing, so the page stops before row 2 with no rows rendered. The
    // notice must name row 2 itself; a plain continue hint would point back at the same row forever.
    let text: string = await impl.readWorkbookRange(CHAT_ID, "WIDE_XLSX", fileId, "Wide", "1:3");
    expect(text.split("\n")).toEqual([
      `Sheet "Wide" rows 1–3 of 3, columns A–AD (blank rows omitted)`,
      `… row 2 alone is over 24 KiB. Read fewer columns of it, e.g. "A2:J2".`,
    ]);
    // A narrower read of that row succeeds and continues past it.
    let narrow: string = await impl.readWorkbookRange(CHAT_ID, "WIDE_XLSX", fileId, "Wide", "A2:J3");
    expect(narrow.split("\n")[1].startsWith("2 A=0-yyy")).toBe(true);
    expect(narrow.split("\n").at(-1)).toBe("3 A=after");
  }, oversizedRowWorkbookUpload()));

  it("reads through the same paging getRows uses, across a chunk seam", () => withWorkbook(
      async ({ impl, fileId, meta, session }) => {
    let boundary = meta.sheets[1].chunkRowStarts[1];
    let shared = readSheetRows(impl, fileId, 1, meta, boundary - 2, boundary + 2);
    expect(shared).toEqual(await session.getRows("Ledger", boundary - 2, boundary + 2));
    expect(shared.map(row => row[0])).toEqual(
        [boundary - 2, boundary - 1, boundary, boundary + 1]);

    let page: string = await impl.readWorkbookRange(
        CHAT_ID, "QUARTERLY_XLSX", fileId, "Ledger", `${boundary}:${boundary + 1}`);
    expect(page.split("\n").slice(1).map(line => line.split(" ").slice(0, 2).join(" ")))
        .toEqual([`${boundary} A=${boundary - 1}`, `${boundary + 1} A=${boundary}`]);
  }));

  it("names the sheets a wrong sheet name missed, and the problem with a range", () =>
      withWorkbook(async ({ impl, fileId }) => {
    await expect(impl.readWorkbookRange(CHAT_ID, "QUARTERLY_XLSX", fileId, "Sales", undefined))
        .rejects.toThrow(
            `This workbook has no sheet named "Sales". Its sheets are: "Summary", "Ledger".`);
    await expect(impl.readWorkbookRange(CHAT_ID, "QUARTERLY_XLSX", fileId, "Summary", "5:6"))
        .rejects.toThrow("Row 5 is past the end of this sheet, which has 3 rows.");
  }));

  it("reads only for the chat that owns the attachment", () => withWorkbook(
      async ({ impl, fileId }) => {
    await expect(impl.readWorkbookRange(OTHER_CHAT_ID, "QUARTERLY_XLSX", fileId, "Summary",
        undefined)).rejects.toThrow("The attachment behind QUARTERLY_XLSX is no longer available.");
    impl.deleteChatAttachmentContent(fileId);
    await expect(impl.readWorkbookRange(CHAT_ID, "QUARTERLY_XLSX", fileId, "Summary",
        undefined)).rejects.toThrow("The attachment behind QUARTERLY_XLSX is no longer available.");
  }));
});
