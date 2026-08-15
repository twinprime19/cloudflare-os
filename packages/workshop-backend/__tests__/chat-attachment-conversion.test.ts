import { describe, expect, it, vi } from "vitest";
import {
  assertConvertedAttachmentTotalWithinBudget,
  validateChatAttachmentUpload,
} from "../src/chat-attachment-validation.js";
import type { DocToMarkdownEnv } from "../src/doc-to-markdown.js";

const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

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
  it("converts DOCX and XLSX for every provider", async () => {
    for (const provider of ["cloudflare", "ollama", "anthropic", "openai", "google"] as const) {
      for (const mimeType of [DOCX_MIME_TYPE, XLSX_MIME_TYPE]) {
        const { toMarkdown, getEnv } = makeConversionEnv();

        const result = await validateChatAttachmentUpload(
          { mimeType, content: ooxmlBytes(), name: "quarterly.docx" },
          provider,
          getEnv,
        );

        expect(toMarkdown).toHaveBeenCalledTimes(1);
        expect(result.convertedFrom).toBe(mimeType);
        expect(result.attachment.mimeType).toBe("text/markdown");
        expect(result.attachment.name).toBe("quarterly.docx");
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

  it("truncates converted output above the per-document cap", async () => {
    const { getEnv } = makeConversionEnv("x".repeat(200 * 1024));

    const result = await validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "long.docx" },
      "cloudflare",
      getEnv,
    );

    expect(result.attachment.content.byteLength).toBe(150 * 1024);
    expect(new TextDecoder().decode(result.attachment.content)).toContain("too long to include in full");
  });

  it("keeps multi-byte characters intact when truncating", async () => {
    // Cutting mid-character would leave replacement characters at the seam.
    const { getEnv } = makeConversionEnv("â".repeat(200 * 1024));

    const result = await validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "long.docx" },
      "cloudflare",
      getEnv,
    );

    const text = new TextDecoder("utf-8", { fatal: false }).decode(result.attachment.content);
    expect(text).not.toContain("�");
  });

  it("describes images embedded in uploaded documents", async () => {
    // The cost boundary: uploads pay for image description, and web-fetch must not (see
    // web-fetch.test.ts for the other half of this pair).
    const { toMarkdown, getEnv } = makeConversionEnv();

    await validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "charts.docx" },
      "cloudflare",
      getEnv,
    );

    const options = toMarkdown.mock.calls[0][1];
    expect(options.conversionOptions.docx.images).toEqual({ convert: true, maxConvertedImages: 20 });
    expect(options.conversionOptions.pdf.images).toEqual({ convert: true, maxConvertedImages: 20 });
  });

  it("refuses conversion when no conversion environment is available", async () => {
    await expect(validateChatAttachmentUpload(
      { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "notes.docx" },
      "cloudflare",
    )).rejects.toThrow("Documents cannot be read in this context.");
  });
});

describe("assertConvertedAttachmentTotalWithinBudget", () => {
  it("accepts a message at the budget and rejects one above it", () => {
    expect(() => assertConvertedAttachmentTotalWithinBudget(400 * 1024)).not.toThrow();
    expect(() => assertConvertedAttachmentTotalWithinBudget(400 * 1024 + 1))
      .toThrow("Too much document text in one message.");
  });
});
