// Chat-attachment limits as the Overseer enforces them: how many attachments one message may
// carry, how many bytes they may total, the canonicalization result that carries the converted
// documents' byte count back to the caller, and the send-time budget both chat entry points apply
// before anything is written.
//
// The rules these build on live in chat-attachment-validation.ts and are covered by
// chat-attachment-validation.test.ts and chat-attachment-conversion.test.ts. What is pinned here
// is the wiring at the Overseer: the limits it declares, and the fact that newChat() and
// sendChatMessage() weigh a message against its converted documents before creating anything.
//
// Runs the real OverseerImpl inside workerd over real Durable Object storage -- the harness
// chat-changes.test.ts uses -- so the limits are exercised through the very methods the RPC
// surface calls. Attachments are staged straight into storage the way uploadChatAttachment()
// writes them, because the upload path needs a Workers AI binding this config does not have.

import { describe, expect, it } from "vitest";
import { env } from "cloudflare:workers";
import { runInDurableObject } from "cloudflare:test";
import type {
  AiChatAuthorInfo, ChatAttachmentHandle, ChatAttachmentRef,
} from "@gadgets/workshop-shared/api";
import { validateChatAttachmentUpload } from "../src/chat-attachment-validation.js";
import type { OverseerDurableObject } from "../src/overseer.js";

declare module "cloudflare:workers" {
  interface ProvidedEnv {
    TEST_OVERSEER: DurableObjectNamespace<OverseerDurableObject>;
  }
}

// 1.75 MiB per stored file (bounded by the Durable Object row limit), 8 MiB across one message,
// and 2 MiB for the message text plus the Markdown extracted from its documents combined.
const PER_FILE_LIMIT = 1792 * 1024;
const PER_MESSAGE_TOTAL_LIMIT = 8 * 1024 * 1024;
const MESSAGE_BUDGET = 2 * 1024 * 1024;

const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const USER: AiChatAuthorInfo = { type: "user", id: "alice@example.com", name: "Alice" };
// newChat()/sendChatMessage() read nothing from the client user stub but its id, for analytics.
const CLIENT_USER = { id: { toString: () => "alice-user-do" } };

// Distinguishable from "absent": both conversion-env getters read the binding and the gateway
// settings off the environment, and the unit test worker binds neither.
const WORKERS_AI_SENTINEL = { gateway: () => ({}) } as unknown as Ai;
const GATEWAY_ENV = {
  WORKERS_AI: WORKERS_AI_SENTINEL,
  CF_AI_GATEWAY: "platform-gateway",
  CF_AI_GATEWAY_ACCOUNT_ID: "gateway-account-id",
  CF_AI_GATEWAY_API_TOKEN: "gateway-token",
};

let doCounter = 0;
async function withImpl(fn: (impl: any) => Promise<void>): Promise<void> {
  let stub = env.TEST_OVERSEER.getByName(`chat-attachment-caps-${++doCounter}`);
  await runInDurableObject(stub, async (instance: OverseerDurableObject) => {
    await fn((instance as unknown as { impl: any }).impl);
  });
}

/** Stages one attachment exactly as uploadChatAttachment() does, and returns its handle. */
function stage(impl: any, options: {
  bytes: number;
  mimeType?: string;
  name?: string;
  /** Set only for uploads converted to Markdown on the way in. */
  convertedFrom?: string;
}): ChatAttachmentHandle {
  let id = crypto.randomUUID();
  impl.storage.chatAttachmentContent.put({
    fileId: id,
    data: new Uint8Array(options.bytes),
    state: {
      type: "staged",
      uploadedAt: Date.now(),
      mimeType: options.mimeType ?? "text/plain",
      name: options.name,
      convertedFrom: options.convertedFrom,
    },
  });
  return { id };
}

/** Four staged files at the per-file limit -- 7 MiB, one MiB short of the per-message total. */
function stageFourFullFiles(impl: any): ChatAttachmentHandle[] {
  return Array.from({ length: 4 }, () => stage(impl, { bytes: PER_FILE_LIMIT }));
}

// Minimal bytes that clear the pre-conversion content checks: the ZIP local-file header every
// OOXML package starts with, followed by the manifest entry that identifies it as OOXML.
function ooxmlBytes(): Uint8Array {
  let header = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);
  let marker = new TextEncoder().encode("[Content_Types].xml");
  let bytes = new Uint8Array(header.length + marker.length);
  bytes.set(header);
  bytes.set(marker, header.length);
  return bytes;
}

describe("canonicalizeChatAttachmentRefs", () => {
  it("accepts eight attachments on one message", () => withImpl(async impl => {
    let handles = Array.from({ length: 8 }, () => stage(impl, { bytes: 16 }));

    let canonical = impl.canonicalizeChatAttachmentRefs(handles, undefined);

    expect(canonical.refs).toHaveLength(8);
    expect(canonical.refs.map((ref: ChatAttachmentRef) => ref.id).toSorted())
        .toEqual(handles.map(handle => handle.id).toSorted());
  }));

  it("rejects a ninth attachment, and names eight as the limit", () => withImpl(async impl => {
    let handles = Array.from({ length: 9 }, () => stage(impl, { bytes: 16 }));

    expect(() => impl.canonicalizeChatAttachmentRefs(handles, undefined))
        .toThrow("You can attach up to 8 attachments.");
  }));

  // Five files either way, so the count limit never decides the outcome and the total does.
  it("accepts attachments totalling 8 MiB across one message, and rejects one byte more",
      () => withImpl(async impl => {
    let full = stageFourFullFiles(impl);
    let remainder = PER_MESSAGE_TOTAL_LIMIT - 4 * PER_FILE_LIMIT;
    let atLimit = [...full, stage(impl, { bytes: remainder })];
    let overLimit = [...full, stage(impl, { bytes: remainder + 1 })];

    let canonical = impl.canonicalizeChatAttachmentRefs(atLimit, undefined);

    expect(canonical.refs).toHaveLength(5);
    expect(canonical.refs.reduce((sum: number, ref: ChatAttachmentRef) => sum + ref.size, 0))
        .toBe(PER_MESSAGE_TOTAL_LIMIT);
    expect(() => impl.canonicalizeChatAttachmentRefs(overLimit, undefined))
        .toThrow("Attached files are too large.");
  }));

  // The converted byte total is only knowable here, while the attachment bytes are in hand, so it
  // travels back with the refs rather than being recomputed by the caller.
  it("returns the refs alongside the converted byte total, not a bare array",
      () => withImpl(async impl => {
    let handle = stage(impl, {
      bytes: 32, mimeType: "text/markdown", name: "report.md", convertedFrom: DOCX_MIME_TYPE,
    });

    let canonical = impl.canonicalizeChatAttachmentRefs([handle], "google");

    expect(Array.isArray(canonical)).toBe(false);
    expect(canonical).toEqual({
      refs: [{ id: handle.id, mimeType: "text/markdown", name: "report.md", size: 32 }],
      convertedBytes: 32,
    });
  }));

  // Markdown a user wrote and attached is stored as-is and bounded by the storage caps; only text
  // extracted from a document counts against the message's context budget.
  it("counts only converted documents toward the converted byte total",
      () => withImpl(async impl => {
    let converted = stage(impl, {
      bytes: 1500, mimeType: "text/markdown", name: "from-docx.md", convertedFrom: DOCX_MIME_TYPE,
    });
    let written = stage(impl, { bytes: 900, mimeType: "text/markdown", name: "notes.md" });

    let canonical = impl.canonicalizeChatAttachmentRefs([converted, written], undefined);

    expect(canonical.refs).toHaveLength(2);
    expect(canonical.convertedBytes).toBe(1500);
  }));
});

describe("combined message and converted-document budget", () => {
  // Message text plus converted document text, over the 2 MiB combined budget.
  const MESSAGE_TEXT = "a".repeat(1024 * 1024);
  const CONVERTED_BYTES = 1536 * 1024;

  // The same message and the same number of bytes twice over: attached as Markdown the user wrote,
  // the chat is created; attached as text extracted from a document, it is refused. Only the
  // second kind is replayed into every model request alongside the message.
  it("weighs converted document text against the message but not bytes stored as-is, " +
      "and creates nothing when it refuses", () => withImpl(async impl => {
    expect(MESSAGE_TEXT.length + CONVERTED_BYTES).toBeGreaterThan(MESSAGE_BUDGET);
    let written = stage(impl, {
      bytes: CONVERTED_BYTES, mimeType: "text/markdown", name: "notes.md",
    });
    let converted = stage(impl, {
      bytes: CONVERTED_BYTES, mimeType: "text/markdown", name: "report.md",
      convertedFrom: DOCX_MIME_TYPE,
    });

    await impl.newChat(CLIENT_USER, { profile: USER }, MESSAGE_TEXT, undefined, [written]);

    expect([...impl.storage.chatMeta.list()]).toHaveLength(1);
    expect(impl.storage.chatAttachmentContent.get(written.id).state.type).toBe("committed");

    await expect(impl.newChat(CLIENT_USER, { profile: USER }, MESSAGE_TEXT, undefined, [converted]))
        .rejects.toThrow("Message is too big.");

    // No second chat, and the attachment is still staged: the refusal came before any write.
    expect([...impl.storage.chatMeta.list()]).toHaveLength(1);
    expect(impl.storage.chatAttachmentContent.get(converted.id).state.type).toBe("staged");
  }));

  it("refuses an over-budget message on an existing chat before the draft is touched",
      () => withImpl(async impl => {
    impl.storage.chatMeta.put(
        { id: 1, title: "Chat", started: new Date(0), lastActive: new Date(0) });
    let converted = stage(impl, {
      bytes: CONVERTED_BYTES, mimeType: "text/markdown", name: "report.md",
      convertedFrom: DOCX_MIME_TYPE,
    });

    await expect(impl.sendChatMessage(
        CLIENT_USER, { profile: USER }, 1, MESSAGE_TEXT, undefined, [converted]))
        .rejects.toThrow("Message is too big.");

    expect([...impl.storage.chats.list()]).toEqual([]);
    expect(impl.storage.chatAttachmentContent.get(converted.id).state.type).toBe("staged");
  }));
});

describe("document conversion environment", () => {
  it("hands upload conversion the same Workers AI environment web fetch gets",
      () => withImpl(async impl => {
    impl.env = GATEWAY_ENV;

    let conversionEnv = impl.getDocumentConversionEnv();
    let webFetchEnv = impl.getWebFetchEnv();

    expect(Object.keys(conversionEnv).toSorted()).toEqual(["ai", "gateway"]);
    expect(conversionEnv.ai).toBe(WORKERS_AI_SENTINEL);
    expect(webFetchEnv.ai).toBe(WORKERS_AI_SENTINEL);
    expect(conversionEnv.gateway).toMatchObject({
      gateway: "platform-gateway",
      accountId: "gateway-account-id",
      apiToken: "gateway-token",
      binding: WORKERS_AI_SENTINEL,
    });
    expect(conversionEnv.gateway).toEqual(webFetchEnv.gateway);
  }));

  // Converting sends the document's bytes to Workers AI, which is exactly what a workspace that
  // has observed sensitive data must not do -- the same reason web fetch is refused there.
  it("refuses to convert an upload in a workspace under sharing lockdown",
      () => withImpl(async impl => {
    impl.storage.containsRestrictedData.put(true);

    expect(() => impl.getDocumentConversionEnv())
        .toThrow("prohibited from converting uploaded documents");
    await expect(validateChatAttachmentUpload(
        { mimeType: DOCX_MIME_TYPE, content: ooxmlBytes(), name: "quarterly.docx" },
        "google",
        () => impl.getDocumentConversionEnv(),
    )).rejects.toThrow("prohibited from converting uploaded documents");
  }));
});
