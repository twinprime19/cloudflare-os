// Parsing a spreadsheet in a dynamic worker of its own (workbook-parser-isolate.ts), through the
// same LOADER binding the workspace uses.
//
// What is pinned: the isolated parse produces byte-for-byte what the parser produces in this
// isolate, for every spreadsheet format; the parser's refusals arrive verbatim; the runtime
// stopping the parser reads as "too large"; and anything else that goes wrong -- a runtime that
// fails to load, a cursor that sends the wrong chunk -- reads as unreadable, never as "too large".
//
// Local workerd does not enforce a dynamic worker's memory limit, so the runtime's exhaustion is
// simulated with a parser that throws the runtime's wording; the real limit is checked on staging.

import { afterEach, describe, expect, it, vi } from "vitest";
import { env } from "cloudflare:workers";
import * as XLSX from "xlsx";
import {
  WORKBOOK_EMPTY_MESSAGE,
  WORKBOOK_TOO_LARGE_MESSAGE,
  WORKBOOK_UNREADABLE_MESSAGE,
  workbookRowChunks,
} from "../src/chat-attachment-workbook.js";
import type { StreamedWorkbook, WorkbookRowChunk } from "../src/chat-attachment-workbook.js";
import { WORKBOOK_PARSE_CPU_MS, parseWorkbookIsolated } from "../src/workbook-parser-isolate.js";
import { parseWorkbookAttachment } from "../src/workbook-parse.js";
import {
  ODS_MIME_TYPE,
  XLSB_MIME_TYPE,
  XLSM_MIME_TYPE,
  XLSX_MIME_TYPE,
  XLS_MIME_TYPE,
} from "../src/workbook-names.js";
import { repeatedCellsOds } from "./repeated-cells-ods.js";

afterEach(() => {
  vi.restoreAllMocks();
});

// A workbook with every cell type the parser renders, a second sheet whose rows need several
// stored chunks, and an empty sheet between them. Dates are built from local components; see
// chat-attachment-workbook.test.ts for why.
function realWorkbook(bookType: XLSX.BookType): Uint8Array {
  let workbook = XLSX.utils.book_new();
  let summary = XLSX.utils.aoa_to_sheet([
    ["Region", "Units", "Shipped", "Due", "Ratio"],
    ["North", 5, true, new Date(2026, 8, 25), 0.25],
    ["South", 7, false, new Date(2026, 8, 25, 14, 30), 1.5],
    [],
    ["Notes", "Figures in VND thousands"],
  ], { cellDates: true });
  summary["F2"] = { t: "e", v: 0x07 };
  summary["!ref"] = "A1:F5";
  XLSX.utils.book_append_sheet(workbook, summary, "Summary");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([]), "Blank");
  // ~1.1 MB of row JSON: more than one chunk.
  let ledger = Array.from({ length: 1_200 }, (_, index) => [index, "x".repeat(900)]);
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(ledger), "Ledger");
  return new Uint8Array(XLSX.write(workbook, { type: "array", bookType }) as ArrayBuffer);
}

function sheetBytes(rows: unknown[][]): Uint8Array {
  let workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  return new Uint8Array(XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as ArrayBuffer);
}

async function drain(workbook: StreamedWorkbook) {
  return {
    summary: workbook.summary,
    meta: workbook.meta,
    chunks: await Array.fromAsync(workbook.rows),
  };
}

// The LOADER with its parser module swapped for `source`, for failures the real parser cannot be
// made to produce on demand.
function loaderRunning(source: string): WorkerLoader {
  return {
    load: code => env.LOADER.load({ ...code, modules: { "parser.js": source } }),
    get: (name, getCode) => env.LOADER.get(name, getCode),
  };
}

// A parser module whose `parse` runs `body`. `cloudflare:workers` is the one import it has.
function parserModule(body: string): string {
  return `
    import { RpcTarget, WorkerEntrypoint } from "cloudflare:workers";
    class Cursor extends RpcTarget {
      constructor(chunks, fail) { super(); this.chunks = chunks; this.fail = fail; }
      next() {
        if (this.chunks.length === 0 && this.fail) throw new Error(this.fail);
        return this.chunks.shift();
      }
    }
    export default class extends WorkerEntrypoint {
      parse(bytes, mimeType, name) { ${body} }
    }`;
}

// A one-sheet, two-chunk index, for the scripted parsers below.
const TWO_CHUNK_META = {
  sheets: [{ name: "Sheet1", rowCount: 2, colCount: 1, chunkRowStarts: [0, 1] }],
  cellCount: 2,
};
const chunkSource = (rowStart: number, chunkIndex = rowStart) =>
  `{ sheetIndex: 0, chunkIndex: ${chunkIndex}, rowStart: ${rowStart}, ` +
  `bytes: new Uint8Array([91, 93]) }`;

function logEvents(spy: { mock: { calls: unknown[][] } }): (string | undefined)[] {
  return spy.mock.calls.map(call => (call[0] as { event?: string })?.event);
}

describe("parseWorkbookIsolated", () => {
  it("loads a fresh parser with no bindings, no network and a CPU limit", async () => {
    let load = vi.spyOn(env.LOADER, "load");

    let parsed = await parseWorkbookIsolated(
        env.LOADER, sheetBytes([["A"], [1]]), XLSX_MIME_TYPE, "small.xlsx");
    await drain(parsed);

    expect(load).toHaveBeenCalledTimes(1);
    expect(load.mock.calls[0][0]).toMatchObject({
      mainModule: "parser.js",
      env: {},
      globalOutbound: null,
      limits: { cpuMs: WORKBOOK_PARSE_CPU_MS },
    });
  });

  for (let [mimeType, bookType] of [
    [XLSX_MIME_TYPE, "xlsx"],
    [XLSM_MIME_TYPE, "xlsm"],
    [XLSB_MIME_TYPE, "xlsb"],
    [XLS_MIME_TYPE, "biff8"],
    [ODS_MIME_TYPE, "ods"],
  ] as const) {
    it(`streams the same summary, index and chunks as the parser in this isolate (${bookType})`,
        async () => {
      let bytes = realWorkbook(bookType);
      let direct = parseWorkbookAttachment(bytes, mimeType, "ledger");

      let isolated = await drain(
          await parseWorkbookIsolated(env.LOADER, bytes, mimeType, "ledger"));

      expect(isolated.summary).toBe(direct.summary);
      expect(isolated.meta).toEqual(direct.meta);
      expect(isolated.chunks).toEqual([...workbookRowChunks(direct.sheets)]);
      // Legacy .xls caps a cell's text at 255 characters, so its ledger fits one chunk.
      if (bookType !== "biff8") {
        expect(isolated.chunks.filter(chunk => chunk.sheetIndex === 2).length).toBeGreaterThan(1);
      }
    });
  }

  it("passes the parser's own refusals through verbatim, without logging them", async () => {
    let warn = vi.spyOn(console, "warn");
    let error = vi.spyOn(console, "error");
    // One row whose JSON cannot fit a stored chunk.
    let oversized = sheetBytes([Array.from({ length: 40 }, () => "y".repeat(30_000))]);
    let corrupt = sheetBytes([["A"]]).slice(0, 200);

    await expect(parseWorkbookIsolated(env.LOADER, oversized, XLSX_MIME_TYPE, "wide.xlsx"))
        .rejects.toThrow(new Error(WORKBOOK_TOO_LARGE_MESSAGE));
    await expect(parseWorkbookIsolated(env.LOADER, corrupt, XLSX_MIME_TYPE, "broken.xlsx"))
        .rejects.toThrow(new Error(WORKBOOK_UNREADABLE_MESSAGE));
    await expect(parseWorkbookIsolated(env.LOADER, sheetBytes([]), XLSX_MIME_TYPE, "blank.xlsx"))
        .rejects.toThrow(new Error(WORKBOOK_EMPTY_MESSAGE));

    expect(logEvents(warn)).not.toContain("attachment.workbook.isolate_exhausted");
    expect(logEvents(error)).not.toContain("attachment.workbook.parse_failed");
  });

  it("reports a parser that fails to load as unreadable, never as too large", async () => {
    let error = vi.spyOn(console, "error").mockImplementation(() => {});

    await expect(parseWorkbookIsolated(
        loaderRunning(`throw new Error("module failed to evaluate");`),
        sheetBytes([["A"]]), XLSX_MIME_TYPE, "quarterly.xlsx"))
        .rejects.toThrow(new Error(WORKBOOK_UNREADABLE_MESSAGE));

    expect(logEvents(error)).toContain("attachment.workbook.parse_failed");
  });

  it("reports the runtime stopping the parser as too large", async () => {
    let warn = vi.spyOn(console, "warn").mockImplementation(() => {});

    await expect(parseWorkbookIsolated(
        loaderRunning(parserModule(`throw new Error("Worker exceeded memory limit.");`)),
        sheetBytes([["A"]]), XLSX_MIME_TYPE, "quarterly.xlsx"))
        .rejects.toThrow(new Error(WORKBOOK_TOO_LARGE_MESSAGE));

    expect(logEvents(warn)).toContain("attachment.workbook.isolate_exhausted");
  });

  it("reports the parser dying partway through the rows as too large", async () => {
    let warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    let parsed = await parseWorkbookIsolated(
        loaderRunning(parserModule(
            `return { summary: "s", meta: ${JSON.stringify(TWO_CHUNK_META)}, ` +
            `rows: new Cursor([${chunkSource(0)}], "Worker exceeded CPU time limit.") };`)),
        sheetBytes([["A"]]), XLSX_MIME_TYPE, "quarterly.xlsx");

    let received: WorkbookRowChunk[] = [];
    await expect((async () => {
      for await (let chunk of parsed.rows) received.push(chunk);
    })()).rejects.toThrow(new Error(WORKBOOK_TOO_LARGE_MESSAGE));

    expect(received.map(chunk => chunk.rowStart)).toEqual([0]);
    expect(logEvents(warn)).toContain("attachment.workbook.isolate_exhausted");
  });

  it("refuses a chunk the index does not name as unreadable", async () => {
    // Every chunk is stored under the index the binding pages against, so one out of place would
    // leave a workbook that reads wrong.
    let error = vi.spyOn(console, "error").mockImplementation(() => {});
    for (let chunks of [
      [chunkSource(0), chunkSource(5, 1)],    // wrong row start
      [chunkSource(0)],                        // ends early
      [chunkSource(0), chunkSource(1), chunkSource(2)],  // runs past the index
    ]) {
      let parsed = await parseWorkbookIsolated(
          loaderRunning(parserModule(
              `return { summary: "s", meta: ${JSON.stringify(TWO_CHUNK_META)}, ` +
              `rows: new Cursor([${chunks.join(",")}]) };`)),
          sheetBytes([["A"]]), XLSX_MIME_TYPE, "quarterly.xlsx");

      await expect(Array.fromAsync(parsed.rows))
          .rejects.toThrow(new Error(WORKBOOK_UNREADABLE_MESSAGE));
    }
    expect(logEvents(error).filter(event => event === "attachment.workbook.parse_failed"))
        .toHaveLength(3);
  });

  it("parses a repetition-encoded ODS in the parser worker, not the caller", async () => {
    // The file that motivates the isolate: a few kilobytes that expand to hundreds of thousands
    // of cells here, and to four million at the staging check's full size. Local workerd enforces
    // no memory limit, so what is checked here is that the repetition is materialized in the
    // parser worker -- the caller receives only the summary, the index and one chunk at a time --
    // and staging checks that the full-size file is refused without harming the workspace.
    let load = vi.spyOn(env.LOADER, "load");
    let bytes = repeatedCellsOds(200, 2_000);
    expect(bytes.byteLength).toBeLessThan(4 * 1024);

    let parsed = await drain(
        await parseWorkbookIsolated(env.LOADER, bytes, ODS_MIME_TYPE, "repeat.ods"));

    expect(load).toHaveBeenCalledTimes(1);
    expect(parsed.meta.cellCount).toBe(200 * 2_000);
    expect(parsed.meta.sheets[0]).toMatchObject({ rowCount: 200, colCount: 2_000 });
  });
});
