import { describe, expect, it, vi } from "vitest";
import {
  assertChatMessageWithinBudget,
  validateChatAttachmentUpload,
} from "../src/chat-attachment-validation.js";
import { isToMarkdownSupportedMimeType } from "../src/doc-to-markdown.js";
import type { DocToMarkdownEnv } from "../src/doc-to-markdown.js";

const PDF_MIME_TYPE = "application/pdf";
const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const XLS_MIME_TYPE = "application/vnd.ms-excel";
const XLSM_MIME_TYPE = "application/vnd.ms-excel.sheet.macroenabled.12";
const XLSB_MIME_TYPE = "application/vnd.ms-excel.sheet.binary.macroenabled.12";
const ODT_MIME_TYPE = "application/vnd.oasis.opendocument.text";
const ODS_MIME_TYPE = "application/vnd.oasis.opendocument.spreadsheet";
const PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DOC_MIME_TYPE = "application/msword";

// The ZIP local-file header every ZIP-based document format starts with, and nothing else: bytes
// that look like some archive but carry no marker identifying the format they claim to be.
function zipBytes(totalBytes = 64): Uint8Array {
  const bytes = new Uint8Array(totalBytes);
  bytes.set([0x50, 0x4b, 0x03, 0x04]);
  return bytes;
}

// Minimal bytes that pass the pre-conversion content checks: the ZIP local-file header every
// OOXML package starts with, followed by the manifest entry name that identifies it as OOXML.
function ooxmlBytes(totalBytes = 64): Uint8Array {
  const header = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);
  const marker = new TextEncoder().encode("[Content_Types].xml");
  const bytes = new Uint8Array(Math.max(totalBytes, header.length + marker.length));
  bytes.set(header);
  bytes.set(marker, header.length);
  return bytes;
}

// An OpenDocument package stores its `mimetype` entry first and uncompressed, so the entry name
// followed by the ODF media type appears literally in the raw bytes. Only the media-type prefix
// shared by every ODF format is checked, so one fixture stands in for .odt and .ods alike.
function odfBytes(totalBytes = 64): Uint8Array {
  const header = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);
  const marker = new TextEncoder().encode("mimetypeapplication/vnd.oasis.opendocument.text");
  const bytes = new Uint8Array(Math.max(totalBytes, header.length + marker.length));
  bytes.set(header);
  bytes.set(marker, header.length);
  return bytes;
}

// The OLE2 compound-file header that legacy Excel workbooks start with.
function oleBytes(totalBytes = 64): Uint8Array {
  const bytes = new Uint8Array(totalBytes);
  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  return bytes;
}

function pdfBytes(totalBytes = 64): Uint8Array {
  const bytes = new Uint8Array(totalBytes);
  bytes.set(new TextEncoder().encode("%PDF-1.7"));
  return bytes;
}

type ToMarkdownStub = ReturnType<typeof vi.fn>;

function makeConversionEnv(markdown = "# Converted\n\nBody text.") {
  const toMarkdown: ToMarkdownStub = vi.fn(async (doc: { name: string; blob: Blob }) => ({
    id: "stub-id",
    name: doc.name,
    mimeType: doc.blob.type,
    format: "markdown" as const,
    tokens: 0,
    data: markdown,
  }));
  const env: DocToMarkdownEnv = { ai: { toMarkdown } as unknown as Ai, gateway: null };
  return { toMarkdown, getEnv: () => env };
}

describe("document upload conversion", () => {
  it("converts every office document for every provider", async () => {
    // No provider has a native input for these formats, so the conversion decision cannot depend
    // on the model -- unlike PDF, which has native paths and is covered separately below.
    const documents: readonly (readonly [string, Uint8Array, string])[] = [
      [DOCX_MIME_TYPE, ooxmlBytes(), "quarterly.docx"],
      [XLSX_MIME_TYPE, ooxmlBytes(), "quarterly.xlsx"],
      [XLS_MIME_TYPE, oleBytes(), "quarterly.xls"],
      [XLSM_MIME_TYPE, ooxmlBytes(), "quarterly.xlsm"],
      [XLSB_MIME_TYPE, ooxmlBytes(), "quarterly.xlsb"],
      [ODT_MIME_TYPE, odfBytes(), "quarterly.odt"],
      [ODS_MIME_TYPE, odfBytes(), "quarterly.ods"],
    ];

    for (const provider of ["cloudflare", "ollama", "anthropic", "openai", "google"] as const) {
      for (const [mimeType, content, name] of documents) {
        const { toMarkdown, getEnv } = makeConversionEnv();

        const result = await validateChatAttachmentUpload(
          { mimeType, content: new Uint8Array(content), name },
          provider,
          getEnv,
        );

        expect(toMarkdown).toHaveBeenCalledTimes(1);
        expect(result.convertedFrom).toBe(mimeType);
        expect(result.attachment.mimeType).toBe("text/markdown");
        expect(result.attachment.name).toBe(name);
        expect(new TextDecoder().decode(result.attachment.content)).toContain("# Converted");
      }
    }
  });

  it("converts PDFs only for providers with no native document input", async () => {
    for (const provider of ["cloudflare", "ollama", undefined] as const) {
      const { toMarkdown, getEnv } = makeConversionEnv();

      const result = await validateChatAttachmentUpload(
        { mimeType: "application/pdf", content: pdfBytes(), name: "report.pdf" },
        provider,
        getEnv,
      );

      expect(toMarkdown).toHaveBeenCalledTimes(1);
      expect(result.attachment.mimeType).toBe("text/markdown");
    }
  });

  it("leaves the native PDF path untouched", async () => {
    for (const provider of ["anthropic", "openai", "google"] as const) {
      const { toMarkdown, getEnv } = makeConversionEnv();

      const result = await validateChatAttachmentUpload(
        { mimeType: "application/pdf", content: pdfBytes(), name: "report.pdf" },
        provider,
        getEnv,
      );

      expect(toMarkdown).not.toHaveBeenCalled();
      expect(result.convertedFrom).toBeUndefined();
      expect(result.attachment.mimeType).toBe("application/pdf");
    }
  });

  it("accepts a document larger than the stored-as-is per-file limit", async () => {
    // The conversion gate runs before the stored-as-is size check, which is bounded by Durable
    // Object storage and does not apply to a document whose bytes are discarded after conversion.
    const { toMarkdown, getEnv } = makeConversionEnv();

    const result = await validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(5 * 1024 * 1024), name: "big.docx" },
      "cloudflare",
      getEnv,
    );

    expect(toMarkdown).toHaveBeenCalledTimes(1);
    expect(result.attachment.mimeType).toBe("text/markdown");
  });

  it("rejects documents above the raw size cap before calling the converter", async () => {
    const { toMarkdown, getEnv } = makeConversionEnv();

    await expect(validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(11 * 1024 * 1024), name: "huge.docx" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("Documents must be 10 MB or smaller.");
    expect(toMarkdown).not.toHaveBeenCalled();
  });

  it("rejects a ZIP that is not an OOXML package", async () => {
    const { toMarkdown, getEnv } = makeConversionEnv();
    const zipWithoutManifest = new Uint8Array(64);
    zipWithoutManifest.set([0x50, 0x4b, 0x03, 0x04]);

    await expect(validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: zipWithoutManifest, name: "payload.docx" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("Chat attachment content does not match its MIME type.");
    expect(toMarkdown).not.toHaveBeenCalled();
  });

  it("rejects a document whose bytes do not match its MIME type", async () => {
    const { toMarkdown, getEnv } = makeConversionEnv();

    await expect(validateChatAttachmentUpload(
      { mimeType: "application/pdf", content: new Uint8Array(64), name: "report.pdf" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("Chat attachment content does not match its MIME type.");
    expect(toMarkdown).not.toHaveBeenCalled();
  });

  it("rejects presentations with an actionable message", async () => {
    const { toMarkdown, getEnv } = makeConversionEnv();

    await expect(validateChatAttachmentUpload(
      { mimeType: PPTX_MIME_TYPE, content: ooxmlBytes(), name: "deck.pptx" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("Export the slides to PDF");
    expect(toMarkdown).not.toHaveBeenCalled();
  });

  it("refuses conversion when the workspace forbids it", async () => {
    // Conversion sends the document to Workers AI, so a locked-down workspace has to refuse it.
    await expect(validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "sensitive.docx" },
      "cloudflare",
      () => { throw new Error("This workspace has observed sensitive data."); },
    )).rejects.toThrow("This workspace has observed sensitive data.");
  });

  it("rejects the upload when conversion fails", async () => {
    const toMarkdown = vi.fn(async () => ({
      id: "x",
      name: "broken.docx",
      mimeType: DOCX_MIME_TYPE,
      format: "error" as const,
      error: "unreadable",
    }));
    const env: DocToMarkdownEnv = { ai: { toMarkdown } as unknown as Ai, gateway: null };

    await expect(validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "broken.docx" },
      "cloudflare",
      () => env,
    )).rejects.toThrow("Markdown conversion failed");
  });

  it("rejects the upload when conversion yields no text", async () => {
    const { getEnv } = makeConversionEnv("   \n  ");

    await expect(validateChatAttachmentUpload(
      { mimeType: XLSX_MIME_TYPE, content: ooxmlBytes(), name: "empty.xlsx" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("No readable text");
  });

  it("stores long converted output whole", async () => {
    const markdown = "x".repeat(1024 * 1024);
    const { getEnv } = makeConversionEnv(markdown);

    const result = await validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "long.docx" },
      "cloudflare",
      getEnv,
    );

    expect(result.attachment.content.byteLength).toBe(markdown.length);
    expect(new TextDecoder().decode(result.attachment.content)).toBe(markdown);
  });

  it("rejects a document whose Markdown does not fit one storage record", async () => {
    // The converted Markdown is stored as a single Durable Object key-value record, so output
    // above that ceiling cannot be stored at all -- and nothing truncates it any more.
    const { getEnv } = makeConversionEnv("x".repeat(1792 * 1024 + 1));

    await expect(validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "enormous.docx" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("This document produced too much text to attach.");
  });

  it("never describes images embedded in uploaded documents", async () => {
    // Describing an embedded image is the only part of conversion that spends Workers AI models,
    // and it makes the upload wait on those calls. Text extraction alone is free and fast.
    const { toMarkdown, getEnv } = makeConversionEnv();

    await validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "charts.docx" },
      "cloudflare",
      getEnv,
    );

    const options = toMarkdown.mock.calls[0][1];
    expect(options.conversionOptions.docx.images).toEqual({ convert: false });
    expect(options.conversionOptions.pdf.images).toEqual({ convert: false });
    expect(options.conversionOptions.html.images.convert).toBe(false);
  });

  it("rejects legacy Word documents", async () => {
    // Cloudflare's toMarkdown() has no .doc converter, so the bytes must never reach it.
    const { toMarkdown, getEnv } = makeConversionEnv();

    await expect(validateChatAttachmentUpload(
      { mimeType: DOC_MIME_TYPE, content: oleBytes(), name: "legacy.doc" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("Unsupported file type");
    expect(toMarkdown).not.toHaveBeenCalled();
  });

  it("rejects bytes that are not the legacy workbook container", async () => {
    const { toMarkdown, getEnv } = makeConversionEnv();

    await expect(validateChatAttachmentUpload(
      { mimeType: XLS_MIME_TYPE, content: new Uint8Array(64), name: "book.xls" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("Chat attachment content does not match its MIME type.");
    expect(toMarkdown).not.toHaveBeenCalled();
  });

  it("rejects a ZIP without the OpenDocument marker", async () => {
    const { toMarkdown, getEnv } = makeConversionEnv();

    await expect(validateChatAttachmentUpload(
      { mimeType: ODT_MIME_TYPE, content: zipBytes(), name: "notes.odt" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("Chat attachment content does not match its MIME type.");
    expect(toMarkdown).not.toHaveBeenCalled();
  });

  it("rejects a macro-enabled workbook without the OOXML manifest", async () => {
    const { toMarkdown, getEnv } = makeConversionEnv();

    await expect(validateChatAttachmentUpload(
      { mimeType: XLSM_MIME_TYPE, content: zipBytes(), name: "macros.xlsm" },
      "cloudflare",
      getEnv,
    )).rejects.toThrow("Chat attachment content does not match its MIME type.");
    expect(toMarkdown).not.toHaveBeenCalled();
  });

  it("converts only types toMarkdown accepts", () => {
    // The convertible set is a subset of Cloudflare's supported-format list; a type outside it
    // would be accepted at upload and then fail in the converter.
    const convertible = [
      PDF_MIME_TYPE, DOCX_MIME_TYPE, XLSX_MIME_TYPE, XLS_MIME_TYPE, XLSM_MIME_TYPE,
      XLSB_MIME_TYPE, ODT_MIME_TYPE, ODS_MIME_TYPE,
    ];

    for (const mimeType of convertible) {
      expect(isToMarkdownSupportedMimeType(mimeType)).toBe(true);
    }
  });

  it("refuses conversion when no conversion environment is available", async () => {
    await expect(validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "notes.docx" },
      "cloudflare",
    )).rejects.toThrow("Documents cannot be read in this context.");
  });
});

describe("combined message and converted attachment budget", () => {
  const BUDGET = 2 * 1024 * 1024;
  const TOO_BIG = "Message is too big.";

  it("accepts converted attachments at the budget and rejects them above it", () => {
    expect(() => assertChatMessageWithinBudget("", BUDGET)).not.toThrow();
    expect(() => assertChatMessageWithinBudget("", BUDGET + 1)).toThrow(TOO_BIG);
  });

  it("counts message text against the same budget", () => {
    expect(() => assertChatMessageWithinBudget("x".repeat(BUDGET), 0)).not.toThrow();
    expect(() => assertChatMessageWithinBudget("x".repeat(BUDGET + 1), 0)).toThrow(TOO_BIG);
  });

  it("adds message text to attachment bytes rather than budgeting them apart", () => {
    expect(() => assertChatMessageWithinBudget("x".repeat(1024), BUDGET - 1024)).not.toThrow();
    expect(() => assertChatMessageWithinBudget("x".repeat(1024), BUDGET - 1023)).toThrow(TOO_BIG);
  });

  it("measures message text in UTF-8 bytes, not characters", () => {
    // Three bytes per character: a message that fits the budget counted as characters is a third
    // over it counted as bytes, which is what the model actually has to hold.
    const message = "\u20ac".repeat(Math.floor(BUDGET / 3) + 1);

    expect(message.length).toBeLessThan(BUDGET);
    expect(() => assertChatMessageWithinBudget(message, 0)).toThrow(TOO_BIG);
  });
});
