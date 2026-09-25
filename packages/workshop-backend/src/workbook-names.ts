// What makes an attachment a workbook, and the env name it is bound under.
//
// Kept apart from the parser (chat-attachment-workbook.ts) so the places that only need to know
// which attachments are workbooks and what they are called -- history replay, compaction, the
// Overseer's naming -- do not pull SheetJS in with it.

import { validateBindingName, type ChatAttachmentRef } from "@gadgets/workshop-shared/api";

export const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
export const XLS_MIME_TYPE = "application/vnd.ms-excel";
export const XLSM_MIME_TYPE = "application/vnd.ms-excel.sheet.macroenabled.12";
export const XLSB_MIME_TYPE = "application/vnd.ms-excel.sheet.binary.macroenabled.12";
export const ODS_MIME_TYPE = "application/vnd.oasis.opendocument.spreadsheet";

/**
 * The spreadsheet formats that take the workbook path instead of Markdown conversion. This is the
 * spreadsheet subset of the convertible document types; the remaining convertible types (PDF,
 * DOCX, ODT) have no rows to bind and keep their existing path.
 */
export const SPREADSHEET_MIME_TYPES: ReadonlySet<string> = new Set([
  XLSX_MIME_TYPE,
  XLS_MIME_TYPE,
  XLSM_MIME_TYPE,
  XLSB_MIME_TYPE,
  ODS_MIME_TYPE,
]);

/**
 * The workbook bindings one message's attachments add to the chat's env, in attachment order:
 * each spreadsheet converted on upload (`convertedFrom` is a spreadsheet type) gets a name from
 * workbookBindingName. An attachment from before workbooks were bound has no `convertedFrom` and
 * binds nothing.
 *
 * The names are derived, not stamped on the message, so every walk that needs them -- history
 * replay, compaction's fold, the Overseer's scope and naming scans -- repeats this derivation and
 * must agree with the others: each passes the full set of names taken at that point of its walk
 * (bindings in scope plus names claimed by pending connection requests) and walks messages in log
 * order. `taken` is not modified; the names derived here are reserved in a private copy, so two
 * files of the same name in one message become `report` and `report_2`.
 */
export function deriveWorkbookBindings(
  attachments: readonly ChatAttachmentRef[] | undefined,
  taken: ReadonlySet<string>,
): { name: string; attachmentId: string }[] {
  let bindings: { name: string; attachmentId: string }[] = [];
  let reserved: Set<string> | undefined;
  for (let attachment of attachments ?? []) {
    if (attachment.convertedFrom === undefined ||
        !SPREADSHEET_MIME_TYPES.has(attachment.convertedFrom)) {
      continue;
    }
    reserved ??= new Set(taken);
    let name = workbookBindingName(attachment.name, reserved);
    reserved.add(name);
    bindings.push({ name, attachmentId: attachment.id });
  }
  return bindings;
}

/**
 * The env name a spreadsheet attachment is bound under in the chat's executeCode environment,
 * derived from its file name and the names already in scope.
 *
 * Every turn replays the chat log and rebuilds the binding map from scratch, so this derivation
 * must depend on nothing but the file name and the names taken before it, in message order: a name
 * that drifted between turns would leave the conversation's earlier `env.NAME` references pointing
 * at nothing, or -- worse -- at a different workbook.
 *
 * The file name's extension is deliberately kept (`big.xlsx` becomes `big_xlsx`): two exports of
 * the same report differing only in format are common, and dropping the extension would collide
 * them into `big` and `big_2`, where neither name says which file it is.
 */
export function workbookBindingName(
  fileName: string | undefined,
  taken: ReadonlySet<string>,
): string {
  // Runs of anything a JavaScript identifier cannot hold -- spaces, dots, accented letters --
  // collapse to one underscore, so the name stays readable instead of growing one underscore per
  // byte of a non-Latin file name.
  let sanitized = (fileName ?? "").replace(/[^A-Za-z0-9_]+/g, "_").replace(/^_+|_+$/g, "");
  let base: string;
  if (sanitized === "") {
    // Nothing survived sanitizing (an unnamed attachment, or a name of punctuation only), so there
    // is nothing to prefix -- the prefix is the whole name.
    base = "file";
  } else if (/^[0-9]/.test(sanitized) || !isUsableBindingName(sanitized)) {
    // A leading digit, a reserved word (`class`), or a name that would collide with an
    // Object.prototype member (`constructor`) cannot be a binding name on its own. Prefixing keeps
    // the file recognizable where dropping the offending part would not; the result is always a
    // valid name, being `file_` followed by identifier characters.
    base = `file_${sanitized}`;
  } else {
    base = sanitized;
  }

  let name = base;
  for (let suffix = 2; taken.has(name); suffix++) name = `${base}_${suffix}`;
  return name;
}

function isUsableBindingName(name: string): boolean {
  try {
    validateBindingName(name);
    return true;
  } catch {
    return false;
  }
}
