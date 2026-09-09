# Chat attachments

A chat message can carry up to 8 attachments. Images are sent as image input, text-like files are
stored as they are, and office documents are converted to Markdown at upload time with the original
bytes discarded. Stored attachments live in the workspace's SQLite-backed Durable Object as one
key-value record each, and that store's 2 MB record limit sets the 1.75 MiB per-file ceiling.

## Supported upload types

**Images** (`.jpg`, `.png`, and the other allowed image types) are sent to the model as image input;
one over 1568 px on its long edge or over 1.75 MiB is downscaled and re-encoded in the browser first.

**Text-like files** — any `text/*` type, such as `.txt`, `.md`, `.csv`, `.html`, `.xml` — are stored
as-is, up to 1.75 MiB per file.

**Documents** are converted to Markdown with Cloudflare Workers AI `toMarkdown()`, on the free
conversion path (no Neurons billed). The uploaded file may be up to 10 MiB; only the Markdown is
kept, stored as `text/markdown`, and the original bytes are discarded. Converted Markdown larger than
1.75 MiB is refused — nothing is truncated. Eight types convert:

| Extension | MIME type |
|---|---|
| `.pdf` | `application/pdf` |
| `.docx` | `application/vnd.openxmlformats-officedocument.wordprocessingml.document` |
| `.xlsx` | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` |
| `.xls` | `application/vnd.ms-excel` |
| `.xlsm` | `application/vnd.ms-excel.sheet.macroenabled.12` |
| `.xlsb` | `application/vnd.ms-excel.sheet.binary.macroenabled.12` |
| `.odt` | `application/vnd.oasis.opendocument.text` |
| `.ods` | `application/vnd.oasis.opendocument.spreadsheet` |

PDF is the one conditional case: it is converted when the selected model's provider cannot read PDFs
natively (the `cloudflare` and `ollama` routes) or when no model is selected. Anthropic, OpenAI, and
Google models get the original PDF. Images embedded inside a document are never described — `toMarkdown()` is always called
with image conversion off, so a chart or photo inside a document contributes no text; attach it as an
image file if the model needs to see it.

## Unsupported

- **`.doc`** (`application/msword`) — Cloudflare has no converter for it. Save as `.docx` instead.
- **`.ppt` / `.pptx`** — rejected with a message asking you to export the deck to PDF.
- **`.numbers`** — browsers never send the MIME type Cloudflare's converter accepts. Export to `.xlsx`.
- **Scanned or image-only PDFs** — no text layer and no OCR, so conversion returns "No readable text
  could be extracted from this document."
- **Files whose bytes do not match their declared type** — refused before conversion with "Chat
  attachment content does not match its MIME type." Windows browsers label `.csv` as
  `application/vnd.ms-excel`, which trips this check; upload the file as `.txt` instead.

## Limits

| Limit | Value |
|---|---|
| Attachments per message | 8 |
| Stored file (image, text-like, converted Markdown) | 1.75 MiB |
| Convertible document, raw upload | 10 MiB |
| Message text + all converted attachment text | 2 MiB (about 570k tokens) |

The last row is a context budget checked at send time; over it the send is refused with "Message is
too big. Shorten the message or attach fewer documents." Format support follows [Cloudflare's
Markdown conversion list](https://developers.cloudflare.com/workers-ai/features/markdown-conversion/#supported-formats).
