import { useEffect, useRef, useState } from "react";
import { plural, t } from "@lingui/core/macro";
import type { RpcStub } from "capnweb";
import type { ChatAttachmentHandle, Overseer } from "@gadgets/workshop-shared/api";
import { reportIssue } from "../../../../errorReporting";
import { formatAttachmentSize } from "../../attachmentFormatting";
import {
  CONVERTIBLE_DOCUMENT_MIME_TYPES,
  MAX_CHAT_ATTACHMENT_TOTAL_BYTES,
  prepareChatAttachment,
} from "./prepareChatAttachment";

export const MAX_COMPOSER_ATTACHMENTS = 8;

export type ComposerAttachment = {
  id: string;
  blob: Blob;
  name?: string;
  previewUrl?: string;
  mimeType: string;
  uploadState: "uploading" | "ready" | "error";
  ref?: ChatAttachmentHandle;
  error?: string;
};

type ComposerAttachmentsOptions = {
  getOverseer: () => Promise<RpcStub<Overseer>> | RpcStub<Overseer>;
  modelId: string | null;
  onError: (message: string) => void;
};

const errorMessage = (error: unknown, fallback: string) =>
  error instanceof Error && error.message ? error.message : fallback;

export const useComposerAttachments = ({
  getOverseer,
  modelId,
  onError,
}: ComposerAttachmentsOptions) => {
  const [attachments, setAttachments] = useState<ComposerAttachment[]>([]);
  const attachmentsRef = useRef<ComposerAttachment[]>([]);
  const stagedCleanupRef = useRef(new Map<string, () => void>());
  // Convertible documents upload one at a time. Each one buffers up to 10 MiB and waits on a
  // Workers AI conversion inside a single Durable Object isolate; firing eight of those at once
  // is not a profile the server is sized for.
  const documentUploadQueueRef = useRef<Promise<void>>(Promise.resolve());
  const mountedRef = useRef(true);
  const getOverseerRef = useRef(getOverseer);
  const onErrorRef = useRef(onError);
  getOverseerRef.current = getOverseer;
  onErrorRef.current = onError;

  const updateAttachments = (
    update: (current: ComposerAttachment[]) => ComposerAttachment[],
  ) => {
    const next = update(attachmentsRef.current);
    attachmentsRef.current = next;
    setAttachments(next);
  };

  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      const abandoned = attachmentsRef.current;
      attachmentsRef.current = [];
      for (const attachment of abandoned) {
        if (attachment.previewUrl) URL.revokeObjectURL(attachment.previewUrl);
        stagedCleanupRef.current.get(attachment.id)?.();
      }
      stagedCleanupRef.current.clear();
    };
  }, []);

  const uploadAttachment = async (
    id: string,
    blob: Blob,
    mimeType: string,
    name: string | undefined,
    uploadModelId: string | null,
  ) => {
    try {
      const content = new Uint8Array(await blob.arrayBuffer());
      if (!mountedRef.current || !attachmentsRef.current.some((item) => item.id === id)) return;
      const overseer = await getOverseerRef.current();
      if (!mountedRef.current || !attachmentsRef.current.some((item) => item.id === id)) return;
      const ref = await overseer.uploadChatAttachment({ mimeType, content, name }, uploadModelId);
      const deleteUpload = () => {
        void overseer.deleteChatAttachment(ref.id).catch(() => {
          // Cleanup is best-effort because the owner may already have disposed its Overseer.
        });
      };
      if (!mountedRef.current || !attachmentsRef.current.some((item) => item.id === id)) {
        deleteUpload();
        return;
      }
      stagedCleanupRef.current.set(id, deleteUpload);
      updateAttachments((current) => current.map((item) =>
        item.id === id ? { ...item, uploadState: "ready", ref } : item));
    } catch (error) {
      if (!mountedRef.current || !attachmentsRef.current.some((item) => item.id === id)) return;
      console.error("Failed to upload chat attachment:", error);
      reportIssue("chat.attachment-upload", error);
      const message = errorMessage(error, t`Upload failed`);
      updateAttachments((current) => current.map((item) =>
        item.id === id ? { ...item, uploadState: "error", error: message } : item));
      onErrorRef.current(errorMessage(error, t`Failed to upload attachment`));
    }
  };

  const addFiles = async (files: FileList | File[]) => {
    const candidates = Array.from(files);
    const initialRoom = MAX_COMPOSER_ATTACHMENTS - attachmentsRef.current.length;
    if (initialRoom <= 0) {
      onErrorRef.current(t`You can attach up to ${MAX_COMPOSER_ATTACHMENTS} attachments`);
      return;
    }
    const accepted = candidates.slice(0, initialRoom);
    if (candidates.length > initialRoom) {
      onErrorRef.current(plural(initialRoom, {
        one: "Only the first attachment was attached",
        other: "Only the first # attachments were attached",
      }));
    }

    const prepared = await Promise.allSettled(accepted.map(async (file) => ({
      file,
      ...(await prepareChatAttachment(file)),
    })));
    if (!mountedRef.current) return;

    for (const result of prepared) {
      if (result.status === "rejected") {
        console.error("Failed to process chat attachment:", result.reason);
        onErrorRef.current(errorMessage(result.reason, t`Failed to process attachment`));
        continue;
      }

      const { file, blob, mimeType } = result.value;
      if (attachmentsRef.current.length >= MAX_COMPOSER_ATTACHMENTS) {
        onErrorRef.current(t`You can attach up to ${MAX_COMPOSER_ATTACHMENTS} attachments`);
        continue;
      }
      // The total budget is on stored bytes. A convertible document is stored as the text
      // extracted from it -- a fraction of what was uploaded -- so charging its raw size here
      // would reject documents well inside the limit the server actually enforces.
      const isConvertibleDocument = CONVERTIBLE_DOCUMENT_MIME_TYPES.has(mimeType);
      const totalBytes = attachmentsRef.current.reduce(
        (sum, attachment) => CONVERTIBLE_DOCUMENT_MIME_TYPES.has(attachment.mimeType)
          ? sum
          : sum + attachment.blob.size,
        0,
      );
      if (!isConvertibleDocument && totalBytes + blob.size > MAX_CHAT_ATTACHMENT_TOTAL_BYTES) {
        onErrorRef.current(
          t`Attached files must total ${formatAttachmentSize(MAX_CHAT_ATTACHMENT_TOTAL_BYTES) ?? ""} or less`,
        );
        continue;
      }

      const id = crypto.randomUUID();
      const previewUrl = mimeType.startsWith("image/") ? URL.createObjectURL(blob) : undefined;
      const attachment: ComposerAttachment = {
        id,
        blob,
        mimeType,
        name: file.name || undefined,
        previewUrl,
        uploadState: "uploading",
      };
      updateAttachments((current) => [...current, attachment]);
      const upload = () => uploadAttachment(id, blob, mimeType, file.name || undefined, modelId);
      if (isConvertibleDocument) {
        // Run the next document whether or not the previous one settled cleanly, so one failed
        // upload cannot strand every document queued behind it.
        documentUploadQueueRef.current = documentUploadQueueRef.current.then(upload, upload);
      } else {
        void upload();
      }
    }
  };

  const removeAttachment = (id: string) => {
    const attachment = attachmentsRef.current.find((item) => item.id === id);
    if (attachment?.previewUrl) URL.revokeObjectURL(attachment.previewUrl);
    stagedCleanupRef.current.get(id)?.();
    stagedCleanupRef.current.delete(id);
    updateAttachments((current) => current.filter((item) => item.id !== id));
  };

  const clearSentAttachments = (sent: readonly ComposerAttachment[]) => {
    const sentIds = new Set(sent.map((attachment) => attachment.id));
    for (const attachment of sent) {
      if (attachment.previewUrl) URL.revokeObjectURL(attachment.previewUrl);
      stagedCleanupRef.current.delete(attachment.id);
    }
    updateAttachments((current) => current.filter((attachment) => !sentIds.has(attachment.id)));
  };

  return {
    attachments,
    addFiles,
    clearSentAttachments,
    removeAttachment,
  };
};
