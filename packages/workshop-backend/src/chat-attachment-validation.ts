import { isTextLikeAttachmentMimeType } from "@gadgets/workshop-shared/api";
import type { AiModelConfig, AiModelProvider, ChatAttachmentUpload } from "@gadgets/workshop-shared/api";
import { PDF_MIME_TYPE } from "./chat-attachment-pdf";
import { convertDocumentToMarkdown } from "./doc-to-markdown";
import type { DocToMarkdownEnv } from "./doc-to-markdown";

// Bounds attachment storage and the bytes replayed into model requests.
//
// Ceiling is the SQLite-backed Durable Object storage limit: a key and its value combined cannot
// exceed 2 MB, and each attachment is stored as one key-value record with no chunking. 1.75 MiB
// leaves roughly 165 KB for the key and the record envelope. Raising this further requires moving
// attachment bytes out of Durable Object storage.
const MAX_CHAT_ATTACHMENT_BYTES = 1792 * 1024; // 1.75 MiB

/**
 * Raw size ceiling for a document that will be converted to Markdown rather than stored as-is.
 * The document itself is never stored, so the Durable Object row limit does not apply; this
 * bounds how much the isolate buffers and hands to Workers AI in one call.
 */
export const MAX_CONVERTIBLE_DOCUMENT_BYTES = 10 * 1024 * 1024; // 10 MiB

/**
 * Ceiling on the Markdown produced from one document. Converted text is replayed into every
 * model request until compaction passes the message, so this is a context budget, not a storage
 * one. Output above the cap is truncated with a note.
 */
const MAX_CONVERTED_DOCUMENT_BYTES = 150 * 1024; // 150 KiB

// Ceiling on converted Markdown across all attachments of a single message.
const MAX_CONVERTED_ATTACHMENT_TOTAL_BYTES = 400 * 1024; // 400 KiB

/**
 * Enforce the per-message budget for text extracted from documents, summed at send time.
 *
 * Without it, a message of several large converted documents can exceed the model's context
 * window on its own -- and the message being sent is the one thing compaction can never drop, so
 * the turn would fail mid-stream instead of being rejected up front.
 */
export function assertConvertedAttachmentTotalWithinBudget(totalBytes: number): void {
  if (totalBytes <= MAX_CONVERTED_ATTACHMENT_TOTAL_BYTES) return;
  throw new Error(
    "Too much document text in one message. Send the documents across separate messages.");
}

/**
 * Upper bound on images described per converted document. Each described image costs two Workers
 * AI model calls (object detection, then image-to-text) that the upload request waits on, and
 * every description competes with the document's own text for the per-document byte budget.
 */
const MAX_DESCRIBED_IMAGES_PER_DOCUMENT = 20;

/** MIME type of the Markdown stored in place of a converted document. */
export const CONVERTED_ATTACHMENT_MIME_TYPE = "text/markdown";

const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

// Office formats built on the OOXML ZIP container. Their bytes are checked for the container's
// manifest entry before conversion (see assertOoxmlContainer).
const OOXML_MIME_TYPES = new Set([DOCX_MIME_TYPE, XLSX_MIME_TYPE]);

// Documents accepted for conversion to Markdown. PDF is conditional -- see shouldConvertUpload.
const CONVERTIBLE_MIME_TYPES = new Set([PDF_MIME_TYPE, DOCX_MIME_TYPE, XLSX_MIME_TYPE]);

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

// Magic-number prefixes checked at upload. Like the image signatures, these only stop mislabeled
// uploads at the door; nothing here parses the content. The OOXML entries are the generic ZIP
// local-file header, which every ZIP container shares and which therefore proves nothing beyond
// "this is some ZIP" -- assertOoxmlContainer adds the format-specific check.
const CONTENT_SIGNATURES = new Map<string, readonly (number | null)[]>([
  ...IMAGE_SIGNATURES,
  [PDF_MIME_TYPE, [0x25, 0x50, 0x44, 0x46, 0x2D]],
  [DOCX_MIME_TYPE, [0x50, 0x4B, 0x03, 0x04]],
  [XLSX_MIME_TYPE, [0x50, 0x4B, 0x03, 0x04]],
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

function sanitizeChatAttachmentMimeType(mimeType: string | undefined): string {
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
function shouldConvertUpload(mimeType: string, provider: AiModelProvider | undefined): boolean {
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

const OOXML_MANIFEST_ENTRY = new TextEncoder().encode("[Content_Types].xml");

// Every OOXML package names its content-type manifest in the archive, so its absence means the
// upload is some other ZIP wearing a .docx/.xlsx MIME type. This is a sanity check on mislabeled
// uploads, NOT a security boundary: it does not inspect the archive's structure and does not stop
// a hostile archive. Resource abuse is bounded elsewhere -- by the raw size cap, and by
// toMarkdown() parsing on Workers AI infrastructure rather than in this isolate, where a failure
// is just a rejected upload.
function assertOoxmlContainer(attachment: ChatAttachmentUpload): void {
  if (!OOXML_MIME_TYPES.has(attachment.mimeType)) return;
  if (containsBytes(attachment.content, OOXML_MANIFEST_ENTRY)) return;
  throw new Error("Chat attachment content does not match its MIME type.");
}

const TRUNCATION_NOTE =
  "\n\n[This document was too long to include in full; the rest was left out.]\n";

// Encode Markdown into at most `maxBytes`, appending a note when anything was dropped. Cutting on
// the encoder's own boundary keeps multi-byte characters (and surrogate pairs) intact.
function encodeCappedMarkdown(markdown: string, maxBytes: number): Uint8Array {
  let encoder = new TextEncoder();
  let full = encoder.encode(markdown);
  if (full.byteLength <= maxBytes) return full;

  let note = encoder.encode(TRUNCATION_NOTE);
  let body = new Uint8Array(maxBytes - note.byteLength);
  let { written } = encoder.encodeInto(markdown, body);

  let result = new Uint8Array(written + note.byteLength);
  result.set(body.subarray(0, written));
  result.set(note, written);
  return result;
}

/** The outcome of validating an upload, including whether it was converted on the way in. */
export type ValidatedChatAttachmentUpload = {
  attachment: ChatAttachmentUpload;
  /** Original MIME type, present only when the upload was converted to Markdown. */
  convertedFrom?: string;
};

async function convertUploadToMarkdown(
  attachment: ChatAttachmentUpload,
  getConversionEnv: (() => DocToMarkdownEnv) | undefined,
): Promise<ValidatedChatAttachmentUpload> {
  let convertedFrom = attachment.mimeType;

  if (attachment.content.byteLength > MAX_CONVERTIBLE_DOCUMENT_BYTES) {
    throw new Error("Documents must be 10 MB or smaller.");
  }
  assertContentMatchesMimeType(attachment);
  assertOoxmlContainer(attachment);

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
    // Uploads are user-initiated and bounded, so images inside the document are described rather
    // than dropped -- charts and diagrams would otherwise vanish from the agent's view of it.
    describeImages: true,
    maxConvertedImages: MAX_DESCRIBED_IMAGES_PER_DOCUMENT,
    gatewayMetadata: { tool: "chatAttachment", automated: false },
  });

  if (!markdown.trim()) {
    throw new Error("No readable text could be extracted from this document.");
  }

  attachment.content = encodeCappedMarkdown(markdown, MAX_CONVERTED_DOCUMENT_BYTES);
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
