// @vitest-environment jsdom

import { describe, expect, it } from "vitest";
// `prepareChatAttachment` throws messages built with the `t` macro, which resolves against
// `@lingui/core`'s global instance. This module is a plain function module, not a component tree,
// so there is no provider to mount; importing `@gadgets/i18n` for its module-scope
// `i18n.loadAndActivate(...)` side effect is what the real app does at `main.tsx` too.
import "@gadgets/i18n";
import {
  MAX_CHAT_ATTACHMENT_BYTES,
  MAX_CONVERTIBLE_DOCUMENT_BYTES,
  prepareChatAttachment,
} from "./prepareChatAttachment";

describe("prepareChatAttachment", () => {
  it("keeps a supported non-image attachment unchanged", async () => {
    const file = new File(["report"], "report.txt", { type: "text/plain" });

    await expect(prepareChatAttachment(file)).resolves.toEqual({
      blob: file,
      mimeType: "text/plain",
    });
  });

  it("rejects a non-image attachment above the upload limit", async () => {
    const file = new File(
      [new Uint8Array(MAX_CHAT_ATTACHMENT_BYTES + 1)],
      "large.bin",
      { type: "application/octet-stream" },
    );

    await expect(prepareChatAttachment(file)).rejects.toThrow(
      "Attachments must be 1.8 MB or smaller.",
    );
  });

  it("takes a convertible document up to its own limit and rejects it above", async () => {
    const withinLimit = new File(
      [new Uint8Array(MAX_CHAT_ATTACHMENT_BYTES + 1)],
      "report.pdf",
      { type: "application/pdf" },
    );

    await expect(prepareChatAttachment(withinLimit)).resolves.toEqual({
      blob: withinLimit,
      mimeType: "application/pdf",
    });

    const overLimit = new File(
      [new Uint8Array(MAX_CONVERTIBLE_DOCUMENT_BYTES + 1)],
      "huge.pdf",
      { type: "application/pdf" },
    );

    await expect(prepareChatAttachment(overLimit)).rejects.toThrow(
      "Attachments must be 10.0 MB or smaller.",
    );
  });

  it.each([
    "application/vnd.ms-excel",
    "application/vnd.ms-excel.sheet.macroenabled.12",
    "application/vnd.ms-excel.sheet.binary.macroenabled.12",
    "application/vnd.oasis.opendocument.text",
    "application/vnd.oasis.opendocument.spreadsheet",
  ])("applies the document limit to every convertible office type: %s", async (mimeType) => {
    const file = new File(
      [new Uint8Array(MAX_CHAT_ATTACHMENT_BYTES + 1)],
      "workbook",
      { type: mimeType },
    );

    await expect(prepareChatAttachment(file)).resolves.toEqual({ blob: file, mimeType });
  });

  // The Blob constructor lowercases `type`, so the platform values a picker can hand through --
  // camel-cased on Windows, parameterised on some Linux desktops -- have to be set on the file.
  it.each([
    [
      "application/vnd.ms-excel.sheet.macroEnabled.12",
      "application/vnd.ms-excel.sheet.macroenabled.12",
    ],
    ["application/vnd.ms-excel; charset=binary", "application/vnd.ms-excel"],
  ])("normalizes %s before classifying it", async (rawType, expected) => {
    const file = new File([new Uint8Array(MAX_CHAT_ATTACHMENT_BYTES + 1)], "workbook");
    Object.defineProperty(file, "type", { value: rawType });

    await expect(prepareChatAttachment(file)).resolves.toEqual({ blob: file, mimeType: expected });
  });
});
