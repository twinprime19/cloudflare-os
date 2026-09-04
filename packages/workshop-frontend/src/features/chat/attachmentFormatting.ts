import { t } from "@lingui/core/macro";

export function formatAttachmentSize(size: number | undefined): string | null {
  if (size === undefined) return null;
  if (size < 1024) return t`${size} B`;
  if (size < 1024 * 1024) return t`${(size / 1024).toFixed(1)} KB`;
  return t`${(size / (1024 * 1024)).toFixed(1)} MB`;
}
