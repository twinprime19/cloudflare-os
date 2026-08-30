import type { Message, Usage } from "@earendil-works/pi-ai";
import type { ModelHandle } from "./ai-models.js";

/**
 * An all-zeros pi Usage record, for synthesizing assistant messages that were never actually
 * produced by a live model call (chat-history replay, compaction prompts).
 */
export function zeroUsage(): Usage {
  return {
    input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
  };
}

/**
 * A failed model request, thrown by completeText() and the agent turn loop. pi never throws for
 * provider failures -- it reports them as a final assistant message with stopReason "error" --
 * so this converts that shape back into an exception for callers that expect one (the overseer's
 * turn error triage and the one-shot completion helpers).
 */
export class AgentTurnError extends Error {
  /** HTTP status of the failing request, when the handle observed a response for it. */
  readonly statusCode?: number;

  constructor(message: string, statusCode?: number) {
    super(message);
    this.statusCode = statusCode;
  }
}

/**
 * Best-effort HTTP status extraction for a failed request. pi reports provider failures as
 * error text only, and the provider SDKs' error messages conventionally begin with the status
 * code (e.g. "400 {...}"); failing that, the handle's transport wrapper records the status of a
 * rejected response on ModelHandle.lastResponse. Either is enough for the overseer's triage
 * (report 5xx/unknown, skip expected 4xx). A transport-level failure (no response) yields
 * undefined.
 */
export function httpStatusFromError(errorMessage: string, handle: ModelHandle)
    : number | undefined {
  const match = /^(\d{3})\b/.exec(errorMessage.trim());
  if (match) return Number(match[1]);
  return handle.lastResponse?.status;
}

/** Longest raw excerpt kept from an error body no parser recognized. */
const ERROR_BODY_SNIPPET_CHARS = 300;

function snippet(text: string): string {
  const oneLine = text.replace(/\s+/g, " ").trim();
  return oneLine.length <= ERROR_BODY_SNIPPET_CHARS
      ? oneLine
      : `${oneLine.slice(0, ERROR_BODY_SNIPPET_CHARS)}…`;
}

/**
 * Distills a failed provider response's body into one line naming the actual reason.
 *
 * The provider SDKs only understand `{error: {message}}` bodies. Cloudflare's AI REST endpoint
 * (Workers AI and the unified model catalog) answers in the Cloudflare API envelope instead --
 * `{success: false, errors: [{code, message}]}` -- so the SDK reports such a failure as
 * "<status> status code (no body)" and discards the envelope. The reason we need (e.g. Workers AI
 * 3007 "Request timeout" vs 3008 "Request was aborted", both HTTP 408) lives only in that body.
 * Unrecognized bodies are kept as a trimmed excerpt; an empty body yields undefined.
 */
export function summarizeProviderErrorBody(text: string): string | undefined {
  const trimmed = text.trim();
  if (trimmed === "") return undefined;
  let parsed: unknown;
  try {
    parsed = JSON.parse(trimmed);
  } catch {
    return snippet(trimmed);
  }
  if (typeof parsed !== "object" || parsed === null) return snippet(trimmed);
  const body = parsed as { errors?: unknown; error?: unknown; message?: unknown };

  // Cloudflare API envelope.
  if (Array.isArray(body.errors) && body.errors.length > 0) {
    const entries = body.errors.map(entry => {
      if (typeof entry !== "object" || entry === null) return String(entry);
      const { code, message } = entry as { code?: unknown; message?: unknown };
      return [
        code === undefined ? undefined : String(code),
        typeof message === "string" ? message : undefined,
      ].filter(part => part !== undefined).join(": ");
    }).filter(entry => entry !== "");
    if (entries.length > 0) return snippet(`Cloudflare ${entries.join("; ")}`);
  }

  // OpenAI / Anthropic shape: `error` is an object carrying `message`, or occasionally a string.
  const error = body.error;
  if (typeof error === "string") return snippet(error);
  if (typeof error === "object" && error !== null) {
    const message = (error as { message?: unknown }).message;
    if (typeof message === "string") return snippet(message);
  }
  if (typeof body.message === "string") return snippet(body.message);
  return snippet(trimmed);
}

/**
 * The message for a failed request: pi's error text, extended with what the provider's response
 * body actually said whenever the SDK dropped it (see summarizeProviderErrorBody). Turns
 * "408 status code (no body)" into "408 status code (no body) — Cloudflare 3007: Request
 * timeout" for the chat error and the logs; leaves messages that already carry the detail alone.
 */
export function describeRequestFailure(errorMessage: string, handle: ModelHandle): string {
  const detail = handle.lastResponse?.errorDetail;
  // The detail is whitespace-collapsed (see snippet), so compare the message the same way.
  if (detail === undefined || errorMessage.replace(/\s+/g, " ").includes(detail)) {
    return errorMessage;
  }
  return `${errorMessage} — ${detail}`;
}

/**
 * Run a single non-streaming-style completion against a ModelHandle and return the response
 * text. Used for one-shot calls: title generation, binding naming, compaction summaries, and
 * LanguageModelBinding.run. Always requests thinking off (one-shots should be quick, and none
 * of them benefit from extended thinking; pre-pi, these calls never configured thinking either).
 * Throws AgentTurnError on provider failure, or the abort reason when `signal` fired.
 */
export async function completeText(handle: ModelHandle, args: {
  systemPrompt?: string;
  /** Convenience: wraps into a single user message. Exactly one of `prompt`/`messages` required. */
  prompt?: string;
  messages?: Message[];
  maxTokens?: number;
  signal?: AbortSignal;
}): Promise<string> {
  const messages: Message[] = args.messages ??
      [{ role: "user", content: args.prompt ?? "", timestamp: Date.now() }];
  const stream = await handle.stream(handle.model, {
    systemPrompt: args.systemPrompt,
    messages,
  }, {
    maxTokens: args.maxTokens,
    signal: args.signal,
    thinking: false,
  });
  const message = await stream.result();
  if (message.stopReason === "error" || message.stopReason === "aborted") {
    // Surface a cancellation as the abort reason, like a directly-aborted request would.
    args.signal?.throwIfAborted();
    const errorMessage = message.errorMessage ?? "The model request failed.";
    throw new AgentTurnError(
        describeRequestFailure(errorMessage, handle), httpStatusFromError(errorMessage, handle));
  }
  return message.content
      .filter(block => block.type === "text")
      .map(block => block.text)
      .join("");
}
