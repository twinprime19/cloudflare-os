# Chat attachments

A chat message can carry up to 8 attachments. Images are sent as image input, text-like files are
stored as they are, prose documents are converted to Markdown at upload time with the original bytes
discarded, and spreadsheets are parsed into stored rows plus a small summary. Stored attachments live
in the workspace's SQLite-backed Durable Object as one key-value record each, and that store's 2 MB
record limit sets the 1.75 MiB per-file ceiling.

## Supported upload types

**Images** (`.jpg`, `.png`, and the other allowed image types) are sent to the model as image input;
one over 1568 px on its long edge or over 1.75 MiB is downscaled and re-encoded in the browser first.

**Text-like files** — any `text/*` type, such as `.txt`, `.md`, `.csv`, `.html`, `.xml` — are stored
as-is, up to 1.75 MiB per file.

### Documents

**Documents** are converted to Markdown with Cloudflare Workers AI `toMarkdown()`, on the free
conversion path (no Neurons billed). The uploaded file may be up to 10 MiB; only the Markdown is
kept, stored as `text/markdown`, and the original bytes are discarded. Converted Markdown larger than
1.75 MiB is refused — nothing is truncated. Three types convert:

| Extension | MIME type |
|---|---|
| `.pdf` | `application/pdf` |
| `.docx` | `application/vnd.openxmlformats-officedocument.wordprocessingml.document` |
| `.odt` | `application/vnd.oasis.opendocument.text` |

PDF is the one conditional case: it is converted when the selected model's provider cannot read PDFs
natively (the `cloudflare` and `ollama` routes) or when no model is selected. Anthropic, OpenAI, and
Google models get the original PDF. Images embedded inside a document are never described —
`toMarkdown()` is always called with image conversion off, so a chart or photo inside a document
contributes no text; attach it as an image file if the model needs to see it.

### Spreadsheets

**Spreadsheets** take a different path: they are never flattened into the prompt. Five types are
recognised:

| Extension | MIME type |
|---|---|
| `.xlsx` | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` |
| `.xls` | `application/vnd.ms-excel` |
| `.xlsm` | `application/vnd.ms-excel.sheet.macroenabled.12` |
| `.xlsb` | `application/vnd.ms-excel.sheet.binary.macroenabled.12` |
| `.ods` | `application/vnd.oasis.opendocument.spreadsheet` |

The workbook is parsed with SheetJS **inside the workspace's Durable Object**. There is no Workers AI
call and nothing leaves the isolate, so the path also works in a workspace under sharing lockdown,
where document conversion is refused. The upload may be up to 10 MiB. Two things are kept:

- a **structural summary**, stored as the attachment's `text/markdown` content — this is the only
  part that replays into a model turn;
- **every row of every sheet**, stored separately as JSON pages of at most 1 MiB each. The model
  never receives these; they are read on demand from `executeCode`.

A workbook therefore costs the same handful of tokens per turn whether it holds ten rows or a
hundred thousand.

#### What the model sees each turn

The summary describes where the data is, not what is in it:

- one heading line — `Workbook big.xlsx — 1 sheet, 200,000 cells.`;
- one line per sheet — `## Sheet "Data" (20,000 rows × 10 cols)`, or `## Sheet "Notes" (empty)`;
- then that sheet's **blocks**. Fully blank rows cut a sheet into bands and fully blank columns cut
  each band into blocks, so a title above a table, two tables side by side, and a footnote below
  them stay separate rather than being merged into one rectangle. A block one column wide has no
  header to pair its values against and is rendered as `Text A1: …`; anything wider is
  `Table A3:D8 (5 data rows)` followed by its header row and its data rows as unpadded CSV.

Previews are bounded. A table's data rows are shown whole while there are 20 or fewer; past that the
first 5 rows are shown followed by `… 19,994 more rows`. A single block's line stops at 2,000
characters (always between values, never mid-value). A sheet cut into more than 30 blocks is listed
as ranges only, with no CSV; past 100 sheets only the headings are rendered. The whole summary is
held under **12 KiB** by dropping whole trailing lines and appending
`… summary truncated; the full workbook is still readable.` Everything elided is still reachable
through the binding below.

#### What the agent can do with it

Every spreadsheet attachment is bound into the chat's `executeCode` environment as `env.<name>`. The
name is derived from the file name — runs of characters a JavaScript identifier cannot hold collapse
to one `_`, and the extension is kept — so `big.xlsx` becomes `env.big_xlsx`, and a second `big.xlsx`
in the same chat becomes `env.big_xlsx_2`. The derivation depends only on the file name and the names
already taken, in message order, so it is the same on every replay of the chat.

The message carrying the attachment ends with one sentence naming the binding, and
`describeBinding("big_xlsx")` returns the sheet list with dimensions plus the API:

- `listSheets()` → `{name, rowCount, colCount}[]`, in the workbook's own sheet order. `rowCount` and
  `colCount` are the used extent, with trailing empty rows and columns excluded.
- `getRows(sheet, start?, end?)` → `CellValue[][]` for rows `[start, end)` of one sheet. Rows are
  0-based counting from the top of the spreadsheet, so row 0 is the sheet's first row — the header
  row when the sheet has one; nothing treats any row as special. `start` defaults to 0 and `end` to
  the sheet's `rowCount`.

At most **20,000 rows** come back per call. A wider range throws, naming the paging call to make
instead, rather than returning a truncated answer the agent would go on to compute over. A `start` at
or past the end of the sheet returns `[]` — how a paging loop knows to stop — and an `end` past the
end is clamped. Only the stored pages that overlap the requested range are read, so a page of a
200,000-row sheet costs the same as a page of a small one.

The binding is read-only: a spreadsheet cannot be edited from `executeCode`. A session opens only
for the owning chat's own agent — a Gadget or a user calling the binding loopback directly is
refused, and so is another chat's attachment. Like any binding in `env`, the agent's own code may
pass the stub on to what it calls; what the holder can then read is that one attachment of that one
chat, and nothing else. There is no per-turn scoping, as there is for worktrees: an attachment's rows
are durable and never edited, so a stub kept past the run that minted it can only read the same rows
again.

#### Fidelity

| Cell | What the agent reads |
|---|---|
| Date or time | An ISO string — `"2026-08-06"`, or `"2026-08-06T09:30:00.000Z"` when the cell carries a time. Never an Excel serial number. |
| Formula | The last value the spreadsheet computed and stored. The formula text is not kept and nothing is recalculated. |
| Error | Its spreadsheet text: `#N/A`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!`, `#GETTING_DATA`. |
| Empty | `null`. Trailing empty cells are trimmed, so a row may be shorter than `colCount`. |

Charts, images, pivot caches, styles and number formats are ignored: only cell values are read, and a
number's displayed formatting is not preserved. Attach a chart as an image file if the model needs to
see it.

## Unsupported

- **`.doc`** (`application/msword`) — Cloudflare has no converter for it. Save as `.docx` instead.
- **`.ppt` / `.pptx`** — rejected with a message asking you to export the deck to PDF.
- **`.numbers`** — browsers never send a MIME type this accepts. Export to `.xlsx`.
- **Scanned or image-only PDFs** — no text layer and no OCR, so conversion returns "No readable text
  could be extracted from this document." A workbook whose sheets hold no cells at all is refused
  with the same message.
- **Files whose bytes do not match their declared type** — refused before conversion or parsing with
  "Chat attachment content does not match its MIME type." Windows browsers label `.csv` as
  `application/vnd.ms-excel`, which trips this check; upload the file as `.txt` instead.
- **Spreadsheets too large to hold in the isolate** — refused with "This workbook is too large to
  read. Split it or remove unused sheets." A spreadsheet that cannot be parsed at all is refused with
  "This spreadsheet could not be read."

## Limits

| Limit | Value |
|---|---|
| Attachments per message | 8 |
| Stored file (image, text-like, converted Markdown, workbook summary) | 1.75 MiB |
| Convertible document or spreadsheet, raw upload | 10 MiB |
| Spreadsheet content, measured by inflating the archive | 32 MiB |
| Spreadsheet rows, counted across all its sheets | 250,000 |
| Spreadsheet row data stored, counted across all its sheets | 32 MiB |
| Spreadsheet summary replayed into context | 12 KiB |
| Stored row page, and therefore the largest single row | 1 MiB |
| Rows per `getRows()` call | 20,000 |
| Message text + all converted attachment text | 2 MiB (about 570k tokens) |

The three spreadsheet ceilings are each measured as it is spent, because none of them can be read off
the file. The content cap is measured by inflating the archive in small pieces and counting, the
inflated bytes being dropped as they are counted: a ZIP states its entry sizes but nothing makes the
statement true, and the parser inflates what the entries really hold. The row and row-data ceilings
bound what the file *becomes* — a sheet's stored extent runs to Excel's 1,048,576 rows whatever it
holds, and a spreadsheet stores a repeated string once and references it from every cell that shows
it, so a small legal file can encode to hundreds of megabytes of rows. The read stops one row past the
row ceiling and the row bytes are counted as they are encoded, so a workbook over either is refused
before the memory is spent rather than after. `.xls` is an uncompressed OLE2 container and skips the
content check, its raw size being its size.

A row whose JSON alone exceeds a stored page is refused rather than split, because a row spanning two
pages would break the index every read pages against.

The last row is a context budget checked at send time; over it the send is refused with "Message is
too big. Shorten the message or attach fewer documents." A spreadsheet contributes only its summary
to that budget, never its rows. Document format support follows [Cloudflare's Markdown conversion
list](https://developers.cloudflare.com/workers-ai/features/markdown-conversion/#supported-formats).
