import assert from "node:assert/strict";
import { join } from "node:path";
import { describe, it } from "node:test";
import { readPoFile } from "./po-catalog.ts";
import { buildShards } from "./shard.ts";

const REPO_ROOT = join(import.meta.dirname, "..", "..");
const EN_PO = join(REPO_ROOT, "packages", "i18n", "locales", "en.po");

describe("buildShards over the real catalog", () => {
  const entries = readPoFile(EN_PO).entries;
  const shards = buildShards(entries, 55);

  it("covers every message exactly once, in catalog order", () => {
    const flattened = shards.flatMap((shard) => shard.messages.map((message) => message.msgid));
    assert.deepEqual(
      flattened,
      entries.map((entry) => entry.msgid),
    );
  });

  it("caps every shard at the requested size", () => {
    for (const shard of shards) {
      assert.ok(shard.messages.length >= 1 && shard.messages.length <= 55);
    }
  });

  it("numbers shards sequentially from 01", () => {
    assert.deepEqual(
      shards.map((shard) => shard.shard),
      shards.map((_, i) => String(i + 1).padStart(2, "0")),
    );
  });

  it("strips the comment prefix from placeholder hints", () => {
    const hinted = shards.flatMap((shard) => shard.messages).find((m) => m.hints.length > 0);
    assert.ok(hinted);
    assert.match(hinted.hints[0], /^placeholder /);
  });
});

describe("buildShards edge shapes", () => {
  it("rejects a non-positive size", () => {
    assert.throws(() => buildShards([], 0), /must be positive/);
  });
});
