import { isTextLikeAttachmentMimeType } from "@gadgets/workshop-shared/api";
import type { AiModelConfig, AiModelProvider, ChatAttachmentUpload } from "@gadgets/workshop-shared/api";
import { PDF_MIME_TYPE } from "./chat-attachment-pdf";
// The spreadsheet MIME types live with the workbook parser, which is the module whose behavior
// they select; importing them keeps one definition rather than two lists to hold in sync.
import {
  ODS_MIME_TYPE,
  SPREADSHEET_MIME_TYPES,
  XLSB_MIME_TYPE,
  XLSM_MIME_TYPE,
  XLS_MIME_TYPE,
  XLSX_MIME_TYPE,
  parseWorkbookAttachment,
} from "./chat-attachment-workbook";
import type { ParsedWorkbook } from "./chat-attachment-workbook";
import { convertDocumentToMarkdown } from "./doc-to-markdown";
import type { DocToMarkdownEnv } from "./doc-to-markdown";

/**
 * Bounds attachment storage and the bytes replayed into model requests.
 *
 * Ceiling is the SQLite-backed Durable Object storage limit: a key and its value combined cannot
 * exceed 2 MB, and each attachment is stored as one key-value record with no chunking. 1.75 MiB
 * leaves roughly 165 KB for the key and the record envelope. Raising this further requires moving
 * attachment bytes out of Durable Object storage.
 */
export const MAX_CHAT_ATTACHMENT_BYTES = 1792 * 1024; // 1.75 MiB

/**
 * Ceiling on how many attachments one chat message can carry.
 *
 * Enforced wherever a batch is assembled -- before any bytes are fetched, so an over-long batch is
 * refused rather than downloaded first -- and again where the refs are canonicalized on the way
 * into the message.
 */
export const MAX_CHAT_ATTACHMENTS_PER_MESSAGE = 8;

/**
 * Raw size ceiling for a document that will be converted to Markdown rather than stored as-is.
 * The document itself is never stored, so the Durable Object row limit does not apply; this
 * bounds how much the isolate buffers and hands to Workers AI in one call.
 */
export const MAX_CONVERTIBLE_DOCUMENT_BYTES = 10 * 1024 * 1024; // 10 MiB

/**
 * Ceiling on everything one message puts in front of the model: the message text plus the Markdown
 * extracted from every document attached to it, combined.
 *
 * This is a context budget, not a storage one. Both halves are replayed into every model request
 * until compaction passes the message, and the message being sent is the one thing compaction can
 * never drop, so an oversized turn would fail mid-stream instead of being rejected up front.
 * Attachments stored as-is -- images, Markdown the user wrote -- are bounded separately by the
 * per-file and per-message storage caps.
 */
const MAX_CHAT_MESSAGE_TEXT_BYTES = 2 * 1024 * 1024; // 2 MiB, roughly 570k tokens

/**
 * Enforce that budget at send time. `convertedAttachmentBytes` is the summed byte length of the
 * message's converted attachments; the message text is measured in UTF-8 bytes so both halves are
 * counted the same way whatever the script.
 */
export function assertChatMessageWithinBudget(
  message: string,
  convertedAttachmentBytes: number,
): void {
  let messageBytes = new TextEncoder().encode(message).byteLength;
  if (messageBytes + convertedAttachmentBytes <= MAX_CHAT_MESSAGE_TEXT_BYTES) return;
  throw new Error("Message is too big. Shorten the message or attach fewer documents.");
}

/** MIME type of the Markdown stored in place of a converted document. */
export const CONVERTED_ATTACHMENT_MIME_TYPE = "text/markdown";

const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const ODT_MIME_TYPE = "application/vnd.oasis.opendocument.text";

// Documents accepted for conversion to Markdown: Cloudflare's free toMarkdown() format list minus
// the image types (converting one is an image-to-text model call), minus the text-like types,
// which are stored as they arrive, and minus Apple Numbers, whose accepted MIME type no browser
// sends -- macOS reports `application/x-iwork-numbers-sffnumbers` and Windows reports nothing, so
// an entry for it would never match an upload. Text extraction for everything here costs no model
// usage. PDF is conditional -- see shouldConvertUpload. The frontend keeps a byte-identical copy
// of this set in prepareChatAttachment.ts so the picker applies the same limit before an upload
// starts.
const CONVERTIBLE_MIME_TYPES = new Set([
  PDF_MIME_TYPE,
  DOCX_MIME_TYPE,
  XLSX_MIME_TYPE,
  XLS_MIME_TYPE,
  XLSM_MIME_TYPE,
  XLSB_MIME_TYPE,
  ODT_MIME_TYPE,
  ODS_MIME_TYPE,
]);

// Presentations have no toMarkdown() support, so they get a rejection that says what to do
// instead of the generic unsupported-type error.
const PRESENTATION_MIME_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.presentationml.presentation", // .pptx
  "application/vnd.ms-powerpoint",                                             // .ppt
]);

const IMAGE_SIGNATURES = new Map<string, readonly (number | null)[]>([
  ["image/jpeg", [0xFF, 0xD8, 0xFF]],
  ["image/png", [0x89, 0x50, 0x4E, 0x47]],
  ["image/webp", [
    0x52, 0x49, 0x46, 0x46,
    null, null, null, null,
    0x57, 0x45, 0x42, 0x50,
  ]],
]);

// The ZIP local-file header every ZIP-based document container starts with. Shared by OOXML and
// OpenDocument packages alike, so on its own it proves nothing beyond "this is some ZIP" --
// assertZipContainerMarker adds the format-specific check.
const ZIP_LOCAL_FILE_HEADER = [0x50, 0x4B, 0x03, 0x04];

// The OLE2 compound-file header that legacy Excel workbooks (.xls) start with.
const OLE2_HEADER = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

// Magic-number prefixes checked at upload. Like the image signatures, these only stop mislabeled
// uploads at the door; nothing here parses the content.
const CONTENT_SIGNATURES = new Map<string, readonly (number | null)[]>([
  ...IMAGE_SIGNATURES,
  [PDF_MIME_TYPE, [0x25, 0x50, 0x44, 0x46, 0x2D]],
  [DOCX_MIME_TYPE, ZIP_LOCAL_FILE_HEADER],
  [XLSX_MIME_TYPE, ZIP_LOCAL_FILE_HEADER],
  [XLSM_MIME_TYPE, ZIP_LOCAL_FILE_HEADER],
  [XLSB_MIME_TYPE, ZIP_LOCAL_FILE_HEADER],
  [ODT_MIME_TYPE, ZIP_LOCAL_FILE_HEADER],
  [ODS_MIME_TYPE, ZIP_LOCAL_FILE_HEADER],
  [XLS_MIME_TYPE, OLE2_HEADER],
]);

const isTextOrImageMime = (mimeType: string) =>
  isTextLikeAttachmentMimeType(mimeType) || IMAGE_SIGNATURES.has(mimeType);

const isTextImageOrPdfMime = (mimeType: string) =>
  isTextOrImageMime(mimeType) || mimeType === PDF_MIME_TYPE;

// pi-ai encodes only text and image content parts, so text + images are universal. PDFs ride an
// image part and are bridged to a provider's native document input where one exists: Gemini takes
// application/pdf inline data as-is, and Anthropic/OpenAI payloads are rewritten in flight (see
// chat-attachment-pdf.ts). Workers AI and Ollama chat endpoints have no document input at all.
const ATTACHMENT_SUPPORT_BY_PROVIDER = {
  anthropic: isTextImageOrPdfMime,
  openai: isTextImageOrPdfMime,
  google: isTextImageOrPdfMime,
  cloudflare: isTextOrImageMime,
  ollama: isTextOrImageMime,
} satisfies Record<AiModelProvider, (mimeType: string) => boolean>;

/**
 * Reduce a declared MIME type to the bare type the validation rules are written against: no
 * parameters, no surrounding whitespace, lower case. Anything missing, blank or carrying a line
 * break becomes `application/octet-stream`.
 *
 * Callers that classify a file before its bytes reach `validateChatAttachmentUpload` -- e.g. an
 * import screening a mailbox listing -- must normalize with this first, or they classify a
 * different string than the pipeline will.
 */
export function sanitizeChatAttachmentMimeType(mimeType: string | undefined): string {
  if (!mimeType || /[\r\n]/.test(mimeType)) return "application/octet-stream";
  return mimeType.split(";", 1)[0].trim().toLowerCase() || "application/octet-stream";
}

function sanitizeChatAttachmentName(name: string | undefined): string | undefined {
  if (!name) return undefined;
  let result = name.replace(/[\r\n]/g, " ").slice(0, 255).trim();
  return result || undefined;
}

/** Whether the selected provider can be sent this attachment type at all. */
export function isChatAttachmentSupportedByProvider(
  provider: AiModelProvider | undefined,
  mimeType: string,
): boolean {
  if (!provider) return isTextOrImageMime(mimeType);
  return ATTACHMENT_SUPPORT_BY_PROVIDER[provider](mimeType);
}

/** Reject an attachment type that the selected provider cannot accept. */
export function assertChatAttachmentSupportedByProvider(
  provider: AiModelProvider | undefined,
  mimeType: string,
  byteLength: number,
): void {
  if (byteLength > MAX_CHAT_ATTACHMENT_BYTES) {
    throw new Error("Chat attachment is too large.");
  }

  if (isChatAttachmentSupportedByProvider(provider, mimeType)) return;

  throw new Error("Unsupported file type");
}

/**
 * Whether an upload is converted to Markdown instead of being stored as-is.
 *
 * PDFs are converted only for providers with no native document input; the ones that take a PDF
 * directly keep the higher-fidelity native path (see chat-attachment-pdf.ts). Office documents
 * have no native path anywhere, so they always convert.
 */
export function shouldConvertUpload(
  mimeType: string,
  provider: AiModelProvider | undefined,
): boolean {
  if (!CONVERTIBLE_MIME_TYPES.has(mimeType)) return false;
  if (mimeType !== PDF_MIME_TYPE) return true;
  return !isChatAttachmentSupportedByProvider(provider, PDF_MIME_TYPE);
}

function assertContentMatchesMimeType(attachment: ChatAttachmentUpload): void {
  let signature = CONTENT_SIGNATURES.get(attachment.mimeType);
  if (!signature) return;
  for (let [index, expected] of signature.entries()) {
    if (expected !== null && attachment.content[index] !== expected) {
      throw new Error("Chat attachment content does not match its MIME type.");
    }
  }
}

// Locate a byte sequence anywhere in a buffer.
function containsBytes(haystack: Uint8Array, needle: Uint8Array): boolean {
  outer: for (let start = 0; start + needle.length <= haystack.length; start++) {
    for (let i = 0; i < needle.length; i++) {
      if (haystack[start + i] !== needle[i]) continue outer;
    }
    return true;
  }
  return false;
}

// Every OOXML package names its content-type manifest in the archive.
const OOXML_MANIFEST_MARKER = new TextEncoder().encode("[Content_Types].xml");

// Every OpenDocument package stores its `mimetype` entry first and uncompressed, so the entry name
// is immediately followed by the ODF media type in the raw bytes. Only the prefix every ODF format
// shares is matched, so one marker covers .odt and .ods alike.
const ODF_MEDIA_TYPE_MARKER = new TextEncoder().encode(
  "mimetypeapplication/vnd.oasis.opendocument",
);

// Byte sequence that identifies a ZIP-based package as the format its MIME type claims. A type
// with no entry here is checked against its magic number alone.
const ZIP_CONTAINER_MARKERS = new Map<string, Uint8Array>([
  [DOCX_MIME_TYPE, OOXML_MANIFEST_MARKER],
  [XLSX_MIME_TYPE, OOXML_MANIFEST_MARKER],
  [XLSM_MIME_TYPE, OOXML_MANIFEST_MARKER],
  [XLSB_MIME_TYPE, OOXML_MANIFEST_MARKER],
  [ODT_MIME_TYPE, ODF_MEDIA_TYPE_MARKER],
  [ODS_MIME_TYPE, ODF_MEDIA_TYPE_MARKER],
]);

// A missing marker means the upload is some other ZIP wearing an office MIME type. This is a
// sanity check on mislabeled uploads, NOT a security boundary: it does not inspect the archive's
// structure and does not stop a hostile archive. Resource abuse is bounded elsewhere -- by the raw
// size cap; by toMarkdown() parsing documents on Workers AI infrastructure rather than in this
// isolate; and, for the spreadsheet types, which do parse here, by the workbook module's own
// ceilings on inflated content, rows and row bytes, each measured as it is spent.
function assertZipContainerMarker(attachment: ChatAttachmentUpload): void {
  let marker = ZIP_CONTAINER_MARKERS.get(attachment.mimeType);
  if (!marker) return;
  if (containsBytes(attachment.content, marker)) return;
  throw new Error("Chat attachment content does not match its MIME type.");
}

/** The outcome of validating an upload, including whether it was converted on the way in. */
export type ValidatedChatAttachmentUpload = {
  attachment: ChatAttachmentUpload;
  /** Original MIME type, present only when the upload was converted to Markdown. */
  convertedFrom?: string;
  /**
   * A spreadsheet's index and row pages, present only for uploads that took the workbook path.
   *
   * The attachment itself carries the summary the model reads; the rows here are stored beside it
   * and reached through a binding, so the caller has to persist both together or the summary would
   * describe data nothing can open.
   */
  workbook?: Pick<ParsedWorkbook, "meta" | "sheets">;
};

/**
 * Parse a spreadsheet upload: the attachment becomes the Markdown summary, and the rows travel
 * back separately for the caller to store.
 *
 * Unlike Markdown conversion this needs no Workers AI binding -- parsing happens in this isolate
 * and no byte leaves it -- so it stays available in a workspace under sharing lockdown, which is
 * exactly the workspace whose spreadsheets must not be shipped to a converter. The checks that
 * precede it are the ones conversion runs, in the same order: a document that is never stored is
 * bounded by the raw cap rather than the storage cap, and the magic-number and container checks
 * reject a mislabeled upload before any parser sees it.
 */
function convertWorkbookUpload(attachment: ChatAttachmentUpload): ValidatedChatAttachmentUpload {
  let convertedFrom = attachment.mimeType;

  if (attachment.content.byteLength > MAX_CONVERTIBLE_DOCUMENT_BYTES) {
    throw new Error("Documents must be 10 MB or smaller.");
  }
  assertContentMatchesMimeType(attachment);
  assertZipContainerMarker(attachment);

  let parsed = parseWorkbookAttachment(
      attachment.content, convertedFrom, attachment.name ?? "spreadsheet");

  let encoded = new TextEncoder().encode(parsed.summary);
  // The summary is budgeted far below this by construction (see MAX_WORKBOOK_PROMPT_BYTES); the
  // storage-record ceiling is asserted anyway, so every stored attachment passes the same gate
  // whichever path produced it.
  if (encoded.byteLength > MAX_CHAT_ATTACHMENT_BYTES) {
    throw new Error("This document produced too much text to attach.");
  }

  attachment.content = encoded;
  attachment.mimeType = CONVERTED_ATTACHMENT_MIME_TYPE;
  return {attachment, convertedFrom, workbook: {meta: parsed.meta, sheets: parsed.sheets}};
}

async function convertUploadToMarkdown(
  attachment: ChatAttachmentUpload,
  getConversionEnv: (() => DocToMarkdownEnv) | undefined,
): Promise<ValidatedChatAttachmentUpload> {
  let convertedFrom = attachment.mimeType;

  if (attachment.content.byteLength > MAX_CONVERTIBLE_DOCUMENT_BYTES) {
    throw new Error("Documents must be 10 MB or smaller.");
  }
  assertContentMatchesMimeType(attachment);
  assertZipContainerMarker(attachment);

  if (!getConversionEnv) {
    throw new Error("Documents cannot be read in this context.");
  }
  // Throws when the workspace is locked down: conversion sends the document to Workers AI, which
  // is exactly what a workspace that has observed sensitive data must not do.
  let env = getConversionEnv();

  let markdown = await convertDocumentToMarkdown(env, {
    bytes: attachment.content,
    mimeType: convertedFrom,
    name: attachment.name ?? "document",
    gatewayMetadata: { tool: "chatAttachment", automated: false },
  });

  if (!markdown.trim()) {
    throw new Error("No readable text could be extracted from this document.");
  }

  let encoded = new TextEncoder().encode(markdown);
  // Storage-layer invariant, not a context budget: the Markdown takes the document's place as one
  // Durable Object key-value record and cannot exceed what a record holds. Converted text is
  // stored whole -- nothing is truncated -- so a document that produces more Markdown than fits
  // has to be refused rather than silently shortened.
  if (encoded.byteLength > MAX_CHAT_ATTACHMENT_BYTES) {
    throw new Error("This document produced too much text to attach.");
  }

  attachment.content = encoded;
  attachment.mimeType = CONVERTED_ATTACHMENT_MIME_TYPE;
  return { attachment, convertedFrom };
}

/**
 * Normalize and validate attachment bytes before staging them in chat storage.
 *
 * Documents the selected model cannot read are converted to Markdown here and stored in that form;
 * the original bytes are discarded. The conversion gate deliberately runs before the stored-as-is
 * checks, whose per-file size limit is bounded by Durable Object storage and does not apply to a
 * document that is never stored.
 *
 * Spreadsheets take that gate too but a different path: they are parsed into a summary plus the
 * row pages returned in `workbook`, which the caller must store alongside the attachment.
 */
export async function validateChatAttachmentUpload(
  attachment: ChatAttachmentUpload,
  provider?: AiModelConfig["provider"],
  getConversionEnv?: () => DocToMarkdownEnv,
): Promise<ValidatedChatAttachmentUpload> {
  attachment.name = sanitizeChatAttachmentName(attachment.name);
  attachment.mimeType = sanitizeChatAttachmentMimeType(attachment.mimeType);

  if (PRESENTATION_MIME_TYPES.has(attachment.mimeType)) {
    throw new Error("Presentations cannot be read yet. Export the slides to PDF and upload that.");
  }

  if (shouldConvertUpload(attachment.mimeType, provider)) {
    // Spreadsheets are read here rather than flattened into Markdown: a workbook's rows are
    // stored and reached through a binding, and only its summary goes into the prompt.
    if (SPREADSHEET_MIME_TYPES.has(attachment.mimeType)) return convertWorkbookUpload(attachment);
    return await convertUploadToMarkdown(attachment, getConversionEnv);
  }

  assertChatAttachmentSupportedByProvider(provider, attachment.mimeType, attachment.content.byteLength);
  assertContentMatchesMimeType(attachment);

  return { attachment };
}

/** Whether a MIME type is one of the image encodings Workshop accepts for chat attachments. */
export function isAllowedChatAttachmentImageMimeType(mimeType: string | undefined): boolean {
  return IMAGE_SIGNATURES.has(sanitizeChatAttachmentMimeType(mimeType));
}
