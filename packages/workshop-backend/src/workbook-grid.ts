// The addressed-row encoding: how spreadsheet rows are written out for the model to read.
//
// Every non-empty row becomes one line that starts with its Excel row number, followed by one
// `COLUMN=value` pair per non-empty cell -- `12 C="LNST Dự án" E=2970249.8`. Each value carries its
// own address, so the model never counts delimiters to find which column a value sits in, and a
// blank row costs nothing because the row numbers already say where everything is.
//
// Long runs of adjacent numbers -- a 34-quarter projection, a monthly series -- are the part of a
// sheet nobody reads cell by cell, and the part a model most often miscounts when it tries. With
// `collapse` on, such a run becomes `O..AV=34 nums (14215554 … 0)`: named, sized and located, but
// not enumerated. Text never collapses, so every label, header and note survives.
//
// Pure and deterministic: the same rows always render to the same bytes, because this text is
// replayed into every turn and a drifting prompt would invalidate cached context.

import type { CellValue } from "./chat-attachment-workbook";

/** A run of more than this many adjacent numeric cells collapses to one range entry. */
export const NUMERIC_RUN_COLLAPSE_THRESHOLD = 8;

/** Explains the encoding once, above the first sheet that uses it. */
export const ADDRESSED_ROW_LEGEND =
  "Lines start with the Excel row number; cells are COLUMN=value; text is quoted; a run of more " +
  `than ${NUMERIC_RUN_COLLAPSE_THRESHOLD} numbers is collapsed to FIRST..LAST=n nums ` +
  "(first … last) — read those with readSheet or compute them in executeCode, never by hand.";

// Excel's last column, XFD. A label past it names no cell of any spreadsheet.
const MAX_COLUMN_INDEX = 16_383;

export type RenderRowsOptions = {
  /** Collapse numeric runs longer than NUMERIC_RUN_COLLAPSE_THRESHOLD. */
  collapse: boolean;
  /** Stop before the first line that would take the text past this many UTF-8 bytes. */
  maxBytes?: number;
};

export type RenderedRows = {
  /** The rendered lines, joined with `\n`. */
  text: string;
  /**
   * 0-based sheet index of the last row fully accounted for -- rendered, or blank and so omitted.
   * One less than the first row's index when not even the first line fit.
   */
  renderedThrough: number;
  /** Whether rows were left out because the next line would have crossed `maxBytes`. */
  truncated: boolean;
};

// Column labels are bijective base 26 -- A..Z, then AA -- written here rather than borrowed from
// SheetJS, so rendering a workbook's rows never loads the parser.

/** The A1 letters of a 0-based column index: 0 → `A`, 47 → `AV`. */
export function columnLabel(index: number): string {
  if (!Number.isInteger(index) || index < 0) throw new RangeError(`Not a column index: ${index}`);
  let label = "";
  for (let rest = index + 1; rest > 0; rest = Math.floor((rest - 1) / 26)) {
    label = String.fromCharCode(0x41 + ((rest - 1) % 26)) + label;
  }
  return label;
}

/**
 * The 0-based column index an A1 column label names: `AV` → 47. Case-insensitive; throws on
 * anything that is not a column of a spreadsheet.
 */
export function parseColumnLabel(label: string): number {
  let normalized = label.trim().toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(normalized)) throw new RangeError(`Not a column label: "${label}"`);
  let index = 0;
  for (let letter of normalized) index = index * 26 + (letter.charCodeAt(0) - 0x40);
  index -= 1;
  if (index > MAX_COLUMN_INDEX) throw new RangeError(`Not a column label: "${label}"`);
  return index;
}

/**
 * Render one non-empty cell value.
 *
 * Numbers are `String(n)` -- unformatted, so the model reads the stored value rather than a
 * display rounding. Dates are already ISO strings by the time they reach here. Text is quoted when
 * it holds a space, a quote or an `=` -- the characters that would otherwise make one value read
 * as the end of it or the start of the next pair -- with inner quotes doubled; an empty string is
 * quoted too, so it cannot read as a missing value. A value's own line breaks are flattened first:
 * the text is read line by line, and a cell that could add a line could imitate a row of its own.
 */
export function formatCellValue(value: NonNullable<CellValue>): string {
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "true" : "false";
  let text = value.replace(/[\r\n]+/g, " ");
  if (text !== "" && !/[ "=]/.test(text)) return text;
  return `"${text.replace(/"/g, '""')}"`;
}

/**
 * Render one row as an addressed line, or `null` when the row holds no cells (blank rows are
 * omitted; the row numbers of the lines around them already say where the gap is).
 */
export function renderAddressedRow(
  rowIndex: number,
  cells: readonly CellValue[],
  options: { collapse: boolean },
): string | null {
  let parts: string[] = [];
  let col = 0;
  while (col < cells.length) {
    let value = cells[col];
    if (value === null) {
      col++;
      continue;
    }
    if (typeof value !== "number") {
      parts.push(`${columnLabel(col)}=${formatCellValue(value)}`);
      col++;
      continue;
    }

    // A numeric run is adjacent numeric cells only: an empty cell or a text cell ends it, so a
    // collapsed range always covers every column between its two ends.
    let runEnd = col;
    while (runEnd + 1 < cells.length && typeof cells[runEnd + 1] === "number") runEnd++;
    let length = runEnd - col + 1;
    if (options.collapse && length > NUMERIC_RUN_COLLAPSE_THRESHOLD) {
      let first = formatCellValue(value);
      let last = formatCellValue(cells[runEnd] as number);
      let range = `${columnLabel(col)}..${columnLabel(runEnd)}`;
      parts.push(`${range}=${length} nums (${first} … ${last})`);
    } else {
      for (let index = col; index <= runEnd; index++) {
        parts.push(`${columnLabel(index)}=${formatCellValue(cells[index] as number)}`);
      }
    }
    col = runEnd + 1;
  }
  return parts.length === 0 ? null : `${rowIndex + 1} ${parts.join(" ")}`;
}

/**
 * Render consecutive rows, `rows[0]` being the sheet's row `startRow` (0-based).
 *
 * Rendering stops at the first line that would cross `maxBytes`, so the cost of rendering a sheet
 * is bounded by the budget, not by the sheet: a 240,000-cell sheet asked for 32 KiB renders about
 * 32 KiB of it and reports `truncated`. Lines are kept whole -- a value is never cut in half.
 */
export function renderAddressedRows(
  rows: readonly (readonly CellValue[])[],
  startRow: number,
  options: RenderRowsOptions,
): RenderedRows {
  let encoder = new TextEncoder();
  let maxBytes = options.maxBytes ?? Number.POSITIVE_INFINITY;
  let lines: string[] = [];
  let usedBytes = 0;

  for (let [offset, row] of rows.entries()) {
    let line = renderAddressedRow(startRow + offset, row, options);
    if (line === null) continue;
    // Every line after the first also costs the newline that joins it to the one before.
    let lineBytes = encoder.encode(line).byteLength + (lines.length === 0 ? 0 : 1);
    if (usedBytes + lineBytes > maxBytes) {
      return { text: lines.join("\n"), renderedThrough: startRow + offset - 1, truncated: true };
    }
    lines.push(line);
    usedBytes += lineBytes;
  }

  return { text: lines.join("\n"), renderedThrough: startRow + rows.length - 1, truncated: false };
}
