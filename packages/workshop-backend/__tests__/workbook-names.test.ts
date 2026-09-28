// The env names spreadsheet attachments are bound under. Replay, compaction and the Overseer's
// naming scans each repeat this derivation over the chat log, so it is pinned: a name that moved
// between turns would leave the conversation's `env.NAME` references pointing at the wrong
// workbook, or at nothing.

import { describe, expect, it } from "vitest";
import type { ChatAttachmentRef } from "@gadgets/workshop-shared/api";
import {
  XLSX_MIME_TYPE, deriveWorkbookBindings, workbookBindingName,
} from "../src/workbook-names.js";

describe("workbookBindingName", () => {
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

function ref(overrides: Partial<ChatAttachmentRef>): ChatAttachmentRef {
  return {id: "file-1", mimeType: "text/markdown", name: "report.xlsx", size: 1,
          convertedFrom: XLSX_MIME_TYPE, ...overrides};
}

describe("deriveWorkbookBindings", () => {
  it("names every spreadsheet in attachment order, deduping within the message", () => {
    expect(deriveWorkbookBindings([
      ref({id: "a"}),
      ref({id: "notes", name: "notes.md", convertedFrom: undefined}),
      ref({id: "b"}),
    ], new Set(["report_xlsx"]))).toEqual([
      {name: "report_xlsx_2", attachmentId: "a"},
      {name: "report_xlsx_3", attachmentId: "b"},
    ]);
  });

  it("binds nothing for attachments that are not converted spreadsheets", () => {
    // No convertedFrom: stored as it arrived, or uploaded before workbooks were bound. A PDF
    // converted to Markdown has no rows to bind.
    expect(deriveWorkbookBindings([
      ref({convertedFrom: undefined}),
      ref({convertedFrom: "application/pdf"}),
    ], new Set())).toEqual([]);
    expect(deriveWorkbookBindings(undefined, new Set())).toEqual([]);
  });

  it("leaves the caller's taken set as it was", () => {
    let taken = new Set(["x"]);
    deriveWorkbookBindings([ref({id: "a"}), ref({id: "b"})], taken);
    expect([...taken]).toEqual(["x"]);
  });
});
