// Splits the source catalog into shard JSON files for the translation swarm, in catalog order.
// Each translation agent reads exactly one shard file (plus the conventions doc) and never the
// catalog itself, so no two agents can ever contend over `vi.po`.
//
//   node scripts/i18n-translate/shard.ts --catalog packages/i18n/locales/en.po --out <dir> [--size 55]
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { parseArgs } from "node:util";
import { readPoFile, type PoEntry } from "./po-catalog.ts";

/** One message as a translation agent sees it. */
export interface ShardMessage {
  msgid: string;
  /** `#. placeholder {0}: expr` hints with the `#. ` prefix stripped -- what a placeholder holds. */
  hints: string[];
}

/** One shard file's content. */
export interface Shard {
  /** Two-digit shard number, `"01"`-based; also the file name stem. */
  shard: string;
  messages: ShardMessage[];
}

/** Partitions catalog entries into shards of at most `size` messages, preserving order. */
export function buildShards(entries: PoEntry[], size: number): Shard[] {
  if (size < 1) throw new Error(`Shard size must be positive, got ${size}.`);
  const shards: Shard[] = [];
  for (let start = 0; start < entries.length; start += size) {
    shards.push({
      shard: String(shards.length + 1).padStart(2, "0"),
      messages: entries.slice(start, start + size).map((entry) => ({
        msgid: entry.msgid,
        hints: entry.comments.map((line) => line.replace(/^#\. /, "")),
      })),
    });
  }
  return shards;
}

function main(): void {
  const { values } = parseArgs({
    options: {
      catalog: { type: "string" },
      out: { type: "string" },
      size: { type: "string", default: "55" },
    },
  });
  if (!values.catalog || !values.out) {
    console.error(
      "Usage: node scripts/i18n-translate/shard.ts --catalog <en.po> --out <dir> [--size 55]",
    );
    process.exit(1);
  }
  const entries = readPoFile(values.catalog).entries;
  const shards = buildShards(entries, Number(values.size));
  mkdirSync(values.out, { recursive: true });
  for (const shard of shards) {
    writeFileSync(join(values.out, `shard-${shard.shard}.json`), JSON.stringify(shard, null, 2) + "\n");
  }
  console.log(
    `Wrote ${shards.length} shards (${entries.length} messages) to ${values.out}: ` +
      shards.map((shard) => `${shard.shard}=${shard.messages.length}`).join(" "),
  );
}

if (process.argv[1] === import.meta.filename) main();
