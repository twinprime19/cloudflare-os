// This file declares the type of a workbook binding -- the programmatic view of a spreadsheet
// attached to a chat. Spreadsheets are the one attachment type the model never receives in full:
// the message carries at most 48 KiB of addressed rows (small sheets in full, large ones as an
// outline), while every row stays in storage and is read on demand through readSheet or the API
// below. A workbook therefore costs a bounded slice of the prompt whether it holds ten rows or a
// hundred thousand.
//
// One binding is registered per spreadsheet attachment as the message carrying it replays (see
// agent.ts), and served in executeCode by WorkbookSessionImpl (workbook-session.ts).
//
// The agent's `describeBinding` tool serves the agent-facing section of this file as text
// (workbook-binding.txt is a symlink to this file, shipped as a Text module -- the
// worktree-binding.txt pattern), so everything below the marker is written for the agent as its
// audience.

// Everything below the following line is returned to agents via `describeBinding`.
// ---- BEGIN AGENT API ----

/** One cell's value. An empty cell is `null`. */
export type CellValue = string | number | boolean | null;

/**
 * A workbook binding holds the full contents of a spreadsheet attached to this chat: every row of
 * every sheet, exactly as the file held them when it was attached. The attachment text in the
 * conversation may show only part of a sheet, so read anything you actually need to compute over
 * through this API.
 *
 * To look at rows rather than compute over them, use the `readSheet` tool instead: it returns a
 * range as text with every cell's address (`readSheet(file: "<binding name>", sheet, "A1:J20")`),
 * without writing any code.
 *
 * The data is read-only and never changes; editing the spreadsheet is not possible from here.
 */
export interface Workbook {
  /**
   * The workbook's sheets, in the order the spreadsheet holds them. `rowCount` and `colCount` are
   * each sheet's used extent, with trailing empty rows and columns excluded, so a sheet holding no
   * data at all reports zero rows.
   */
  listSheets(): Promise<{name: string; rowCount: number; colCount: number}[]>;

  /**
   * Rows `[start, end)` of one sheet as arrays of cell values. `sheet` is a sheet name exactly as
   * `listSheets()` reports it; `start` defaults to 0 and `end` to the sheet's `rowCount`, so
   * `getRows("Sales")` returns the whole sheet.
   *
   * Rows are 0-based and count from the top of the spreadsheet: row 0 is the sheet's first row --
   * which is the header row when the sheet has one, since nothing here treats any row as special.
   * (Spreadsheet software numbers the same row 1.)
   *
   * Cell values are strings, numbers, booleans, or `null` for an empty cell. Dates and times
   * arrive as ISO strings (`"2026-08-06"`, or `"2026-08-06T09:30:00.000Z"` for a cell that also
   * carries a time), formulas as their last computed value, and error cells as their spreadsheet
   * text (`"#N/A"`, `"#DIV/0!"`). Trailing empty cells are trimmed, so a row may be shorter than
   * `colCount`; read a missing element as empty rather than assuming every row is full width.
   *
   * At most 20000 rows come back per call. A wider range throws instead of truncating silently, so
   * page through a large sheet with `start`/`end` -- `getRows("Sales", 0, 20000)`, then
   * `getRows("Sales", 20000, 40000)`, and so on. A `start` at or past the end of the sheet returns
   * an empty array, which is how such a loop knows to stop.
   *
   * Prefer computing over printing: 20000 rows returned so you can read them yourself would swamp
   * your context, while summing, filtering or grouping them inside executeCode and printing the
   * answer costs nothing.
   */
  getRows(sheet: string, start?: number, end?: number): Promise<CellValue[][]>;
}
