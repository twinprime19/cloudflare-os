// User avatar validation + storage. The single source of truth for what a valid avatar is, so
// every caller (the browser-driven RPC today, sign-in profile seeding later) agrees.

const MAX_AVATAR_BYTES = 100 * 1024;

/**
 * Env avatar storage needs. Narrowed to the one binding avatars live in (KV, global, keyed by
 * user id -- not the user's DO storage).
 */
export type AvatarEnv = Pick<Cloudflare.Env, 'AVATARS'>;

/**
 * Validates that `data` is an acceptable avatar image: at most 100 KB, and JPEG or PNG by
 * magic-byte header. Throws with a user-facing message otherwise.
 */
export function validateAvatarBytes(data: Uint8Array): void {
  if (data.byteLength > MAX_AVATAR_BYTES) {
    throw new Error("Avatar too large (max 100 KB)");
  }
  // Verify the data starts with a known image magic-byte header.
  let isJpeg = data[0] === 0xFF && data[1] === 0xD8 && data[2] === 0xFF;
  let isPng = data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4E && data[3] === 0x47;
  if (!isJpeg && !isPng) {
    throw new Error("Avatar must be a JPEG or PNG image");
  }
}

/**
 * Writes an avatar to KV, keyed by user id (= email). Callers must validate `data` first (see
 * `validateAvatarBytes`); this function does not validate.
 */
export async function putUserAvatar(env: AvatarEnv, userId: string, data: Uint8Array): Promise<void> {
  await env.AVATARS.put(userId, data);
}

/**
 * Whether an avatar is already stored for this user. This reads the value, not just its presence —
 * the image bytes come back and are discarded, which is cheap because avatars are capped at 100 KB.
 */
export async function userAvatarExists(env: AvatarEnv, userId: string): Promise<boolean> {
  let result = await env.AVATARS.get(userId, "arrayBuffer");
  return result !== null;
}
