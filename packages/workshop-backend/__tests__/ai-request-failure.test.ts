import { describe, expect, it } from "vitest";
import type { AiChatAuthorInfo, AiModelConfig } from "@gadgets/workshop-shared/api";
import { describeRequestFailure, httpStatusFromError, summarizeProviderErrorBody }
  from "../src/ai-invoke.js";
import { getModel, type ModelHandle } from "../src/ai-models.js";

// What a failed model request leaves behind. The provider SDKs drop any error body they don't
// recognize -- Cloudflare's API envelope in particular -- and pi then reports only
// "<status> status code (no body)", so the handle's transport wrapper records the response
// itself and these helpers turn it into the reason shown in the chat and the logs.

const CLOUDFLARE_ENVELOPE = JSON.stringify({
  success: false,
  errors: [{ code: 3007, message: "Request timeout" }],
  messages: [],
  result: null,
});

const INITIATOR: AiChatAuthorInfo = { type: "user", id: "user-123", name: "User" };

// A unified-catalog model: the path every Contoso model takes (api.cloudflare.com/.../ai/v1 over
// pi's OpenAI chat-completions adapter), and the one whose error bodies the SDK drops.
const UNIFIED_CONFIG: AiModelConfig = {
  provider: "cloudflare",
  model: "alibaba/qwen3.7-plus",
  apiToken: "ignored-in-gateway-mode",
};

const GATEWAY_ENV = {
  CF_AI_GATEWAY: "platform-gateway",
  CF_AI_GATEWAY_ACCOUNT_ID: "gateway-account-id",
  CF_AI_GATEWAY_API_TOKEN: "gateway-token",
  CF_AI_GATEWAY_PROVIDERS: "cloudflare",
} as Cloudflare.Env;

function handleWith(errorDetail?: string): ModelHandle {
  return {
    lastResponse: errorDetail === undefined ? undefined : { status: 408, errorDetail },
  } as unknown as ModelHandle;
}

function respondWith(body: string, init: ResponseInit): typeof fetch {
  return (async () => new Response(body, init)) as typeof fetch;
}

// Drives one request through the handle with the given transport and returns pi's error text.
async function failRequest(handle: ModelHandle, fetch: typeof fetch): Promise<string> {
  const stream = await handle.stream(handle.model, {
    messages: [{ role: "user", content: "hello", timestamp: 0 }],
  }, { fetch, maxRetries: 0 });
  const message = await stream.result();
  expect(message.stopReason).toBe("error");
  return message.errorMessage ?? "";
}

describe("summarizeProviderErrorBody", () => {
  it("names the code and message of the Cloudflare API envelope", () => {
    expect(summarizeProviderErrorBody(CLOUDFLARE_ENVELOPE)).toBe("Cloudflare 3007: Request timeout");
    expect(summarizeProviderErrorBody(JSON.stringify({
      success: false,
      errors: [{ code: 3007, message: "Request timeout" }, { code: 3008, message: "Request was aborted" }],
    }))).toBe("Cloudflare 3007: Request timeout; 3008: Request was aborted");
  });

  it("uses error.message for OpenAI- and Anthropic-shaped bodies", () => {
    expect(summarizeProviderErrorBody('{"type":"error","error":{"type":"overloaded_error","message":"Overloaded"}}'))
        .toBe("Overloaded");
    expect(summarizeProviderErrorBody('{"error":{"message":"Rate limit reached","type":"rate_limit"}}'))
        .toBe("Rate limit reached");
    expect(summarizeProviderErrorBody('{"error":"invalid key"}')).toBe("invalid key");
    expect(summarizeProviderErrorBody('{"message":"upstream unavailable"}')).toBe("upstream unavailable");
  });

  it("keeps a one-line excerpt of anything unrecognized", () => {
    expect(summarizeProviderErrorBody("<html>\n  <body>upstream   timeout</body>\n</html>"))
        .toBe("<html> <body>upstream timeout</body> </html>");
    expect(summarizeProviderErrorBody('{"unexpected":true}')).toBe('{"unexpected":true}');
    expect(summarizeProviderErrorBody("x".repeat(400))).toBe(`${"x".repeat(300)}…`);
  });

  it("yields nothing for an empty body", () => {
    expect(summarizeProviderErrorBody("")).toBeUndefined();
    expect(summarizeProviderErrorBody(" \n\t")).toBeUndefined();
  });
});

describe("describeRequestFailure", () => {
  it("appends the recorded reason to a bare SDK status message", () => {
    expect(describeRequestFailure("408 status code (no body)",
        handleWith("Cloudflare 3007: Request timeout")))
        .toBe("408 status code (no body) — Cloudflare 3007: Request timeout");
  });

  it("leaves a message that already carries the reason, or has none recorded, unchanged", () => {
    expect(describeRequestFailure("400 stubbed", handleWith("stubbed"))).toBe("400 stubbed");
    // The SDK may quote a raw body verbatim, line breaks included; the recorded excerpt is
    // whitespace-collapsed, so the two must still be recognized as the same text.
    expect(describeRequestFailure("502 <html>\n  <body>bad gateway</body>\n</html>",
        handleWith("<html> <body>bad gateway</body> </html>")))
        .toBe("502 <html>\n  <body>bad gateway</body>\n</html>");
    expect(describeRequestFailure("connection reset", handleWith())).toBe("connection reset");
  });
});

describe("ModelHandle failed-request capture", () => {
  it("records the Cloudflare envelope the SDK drops on a 408", async () => {
    const handle = getModel(GATEWAY_ENV, UNIFIED_CONFIG, INITIATOR);
    const errorMessage = await failRequest(handle, respondWith(CLOUDFLARE_ENVELOPE, {
      status: 408,
      headers: { "content-type": "application/json", "cf-aig-log-id": "log-408" },
    }));

    // pi's own report carries the status but not the reason.
    expect(errorMessage).toMatch(/^408\b/);
    expect(errorMessage).not.toContain("3007");
    expect(handle.lastResponse).toEqual({
      status: 408,
      aiGatewayLogId: "log-408",
      errorDetail: "Cloudflare 3007: Request timeout",
    });
    expect(httpStatusFromError(errorMessage, handle)).toBe(408);
    expect(describeRequestFailure(errorMessage, handle))
        .toBe(`${errorMessage} — Cloudflare 3007: Request timeout`);
  }, 15000);

  it("records an empty error body as such and replaces the record on the next request", async () => {
    const handle = getModel(GATEWAY_ENV, UNIFIED_CONFIG, INITIATOR);
    await failRequest(handle, respondWith("", { status: 408 }));
    expect(handle.lastResponse).toEqual({ status: 408, aiGatewayLogId: undefined, errorDetail: undefined });

    const errorMessage = await failRequest(handle, respondWith(
        JSON.stringify({ error: { type: "bad_request", message: "stubbed" } }),
        { status: 400, headers: { "content-type": "application/json" } }));
    expect(handle.lastResponse).toEqual({ status: 400, aiGatewayLogId: undefined, errorDetail: "stubbed" });
    // The SDK understood this body, so the message already carries the reason.
    expect(describeRequestFailure(errorMessage, handle)).toBe(errorMessage);
  }, 15000);

  it("rethrows a transport failure, leaving no response recorded", async () => {
    const handle = getModel(GATEWAY_ENV, UNIFIED_CONFIG, INITIATOR);
    const errorMessage = await failRequest(handle, (async () => {
      throw new TypeError("connection reset");
    }) as typeof fetch);
    // The SDK reports a thrown fetch generically; the wrapper's log line carries the cause.
    expect(errorMessage).toMatch(/connection error/i);
    expect(handle.lastResponse).toBeUndefined();
    expect(httpStatusFromError(errorMessage, handle)).toBeUndefined();
  }, 15000);
});
