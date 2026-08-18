import { describe, expect, it } from "vitest";
import { putUserAvatar, userAvatarExists, validateAvatarBytes, type AvatarEnv } from "../src/avatar.js";

const JPEG_HEADER = new Uint8Array([0xFF, 0xD8, 0xFF, 0xE0, 0, 0, 0, 0]);
const PNG_HEADER = new Uint8Array([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

// Minimal in-memory stand-in for the AVATARS KV binding, mirroring the pattern in
// format-blueprints.test.ts.
function makeEnv(): { kv: Map<string, Uint8Array>; env: AvatarEnv } {
  let kv = new Map<string, Uint8Array>();
  return {
    kv,
    env: {
      AVATARS: {
        put: async (key: string, value: Uint8Array) => { kv.set(key, value); },
        get: async (key: string, _type: "arrayBuffer") => kv.get(key)?.buffer ?? null,
        delete: async (key: string) => { kv.delete(key); },
      },
    } as unknown as AvatarEnv,
  };
}

describe("validateAvatarBytes", () => {
  it("accepts a JPEG under the size limit", () => {
    expect(() => validateAvatarBytes(JPEG_HEADER)).not.toThrow();
  });

  it("accepts a PNG under the size limit", () => {
    expect(() => validateAvatarBytes(PNG_HEADER)).not.toThrow();
  });

  it("rejects data over 100 KB with the exact message", () => {
    let oversized = new Uint8Array(100 * 1024 + 1);
    oversized.set(JPEG_HEADER);
    expect(() => validateAvatarBytes(oversized)).toThrow("Avatar too large (max 100 KB)");
  });

  it("rejects non-JPEG/PNG data with the exact message", () => {
    let gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]);
    expect(() => validateAvatarBytes(gif)).toThrow("Avatar must be a JPEG or PNG image");
  });
});

describe("putUserAvatar / userAvatarExists", () => {
  it("writes to KV keyed by user id and reports existence", async () => {
    let { kv, env } = makeEnv();

    expect(await userAvatarExists(env, "user@example.com")).toBe(false);

    await putUserAvatar(env, "user@example.com", JPEG_HEADER);

    expect(kv.get("user@example.com")).toBe(JPEG_HEADER);
    expect(await userAvatarExists(env, "user@example.com")).toBe(true);
    expect(await userAvatarExists(env, "someone-else@example.com")).toBe(false);
  });
});
