// Document-to-Markdown conversion via Cloudflare Workers AI's `env.WORKERS_AI.toMarkdown()`.
//
// Shared by the two call sites that need it: the agent's webFetch tool (converting fetched
// documents) and chat attachment upload (converting uploaded PDF/Office documents). Parsing runs
// on Workers AI infrastructure, not in this isolate.
//
// Embedded-image description is opt-in per call site because it is the one part of conversion
// that spends Workers AI models (object detection followed by image-to-text). webFetch leaves it
// off: it runs automatically against arbitrary third-party URLs, so its per-call cost would be
// uncontrolled. Chat upload turns it on: it is user-initiated and bounded by the upload caps.

import type { AiGatewayConfig } from "./ai-gateway";

/**
 * The bits of the Workers AI binding and gateway config that conversion needs. Kept narrow so
 * callers can pass a stub in tests without constructing a full Cloudflare.Env.
 */
export type DocToMarkdownEnv = {
  ai: Ai;
  gateway: AiGatewayConfig | null;
};

export type DocToMarkdownInput = {
  bytes: Uint8Array;
  /** Base MIME type, without parameters. */
  mimeType: string;
  /** Document name, used by `toMarkdown()` as a format-detection hint. */
  name: string;
  /**
   * Describe images embedded in the document, writing a natural-language description into the
   * markdown. Costs Workers AI model usage, so it defaults to off; see the file header.
   */
  describeImages?: boolean;
  /** Upper bound on described images per document. Only meaningful with `describeImages`. */
  maxConvertedImages?: number;
  /** Origin that relative links in HTML documents resolve against. */
  htmlHostname?: string;
  /** AI Gateway metadata identifying the call site. */
  gatewayMetadata?: GatewayOptions["metadata"];
};

// MIME types that `env.WORKERS_AI.toMarkdown()` can convert. Derived from the public list of
// supported formats:
// https://developers.cloudflare.com/workers-ai/features/markdown-conversion/supported-formats/
//
// Image MIME types are intentionally excluded -- converting an image *is* an image-to-text model
// call, with no free text-extraction path to fall back on.
const TO_MARKDOWN_MIME_TYPES = new Set([
  "text/html",
  "application/xhtml+xml",
  "application/pdf",
  "application/xml",
  "text/xml",
  "text/csv",
  // Office / OpenDocument
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document", // .docx
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",       // .xlsx
  "application/vnd.ms-excel",                                                // .xls
  "application/vnd.ms-excel.sheet.macroenabled.12",                          // .xlsm
  "application/vnd.ms-excel.sheet.binary.macroenabled.12",                   // .xlsb
  "application/vnd.oasis.opendocument.spreadsheet",                          // .ods
  "application/vnd.oasis.opendocument.text",                                 // .odt
  "application/vnd.apple.numbers",                                           // .numbers
]);

/** Whether `toMarkdown()` accepts this MIME type without invoking a paid model. */
export function isToMarkdownSupportedMimeType(mimeType: string): boolean {
  return TO_MARKDOWN_MIME_TYPES.has(mimeType);
}

// `toMarkdown()` uses the Workers AI binding, so only apply the same-account Workers AI gateway
// resolved by AiGatewayConfig. A cross-account platform gateway cannot be used by this binding.
function buildGatewayOptions(
  gateway: AiGatewayConfig | null,
  metadata: GatewayOptions["metadata"],
): GatewayOptions | undefined {
  if (!gateway) return undefined;
  if (!gateway.workersAiGateway) return undefined;
  return { id: gateway.workersAiGateway, metadata };
}

// Image handling is set on every format key that accepts it, so a caller's choice holds whatever
// the document turns out to be. `convertOGImage` is HTML-only.
function buildConversionOptions(input: DocToMarkdownInput): ConversionOptions {
  const images: EmbeddedImageConversionOptions = input.describeImages
    ? { convert: true, maxConvertedImages: input.maxConvertedImages }
    : { convert: false };

  return {
    html: {
      hostname: input.htmlHostname,
      images: { ...images, convertOGImage: input.describeImages === true },
    },
    pdf: { images },
    docx: { images },
  };
}

/**
 * Convert a document to Markdown. Throws with a contextual error if the conversion fails or
 * produces nothing; callers are expected to surface that to the user rather than store a
 * half-converted document.
 */
export async function convertDocumentToMarkdown(
  env: DocToMarkdownEnv,
  input: DocToMarkdownInput,
): Promise<string> {
  const result = await env.ai.toMarkdown(
    {
      name: input.name,
      blob: new Blob([input.bytes], { type: input.mimeType }),
    },
    {
      gateway: buildGatewayOptions(env.gateway, input.gatewayMetadata),
      conversionOptions: buildConversionOptions(input),
    },
  );

  if (result.format === "error") {
    throw new Error(`Markdown conversion failed: ${result.error}`);
  }
  return result.data;
}
