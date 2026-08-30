// Unified-catalog models: third-party models the platform serves under a `provider/model` id
// through Cloudflare's account-level AI REST endpoint, rather than the AI Gateway's workers-ai
// route, which only hosts `@cf/` models.
//
// Two things are pinned. The deployment's suggested Cloudflare catalog is exactly one such model
// with a published context window that other code (compaction budgets above all) computes from.
// And getModel() routes that id to the REST endpoint, refusing up front when the deployment has
// only the Workers AI binding -- a transport that cannot carry that endpoint, and would otherwise
// fail per turn, mid-stream, instead of at configuration time.
//
// Same conventions as ai-models.test.ts: the real pi-ai stack, no module mocks, routing asserted
// on the handle's descriptor and request-level behaviour by driving `stream` with a fetch stub.

import { beforeEach, describe, expect, it } from "vitest";
import type { AiChatAuthorInfo, AiModelConfig } from "@gadgets/workshop-shared/api";
import { SUGGESTED_MODELS, WORKERS_AI_OUTPUT_LIMIT } from "@gadgets/workshop-shared/api";
import { getModel, type ModelHandle } from "../src/ai-models.js";

const INITIATOR: AiChatAuthorInfo = { type: "user", id: "user-123", name: "User" };

const UNIFIED_MODEL_ID = "alibaba/qwen3.7-plus";

const UNIFIED_CONFIG: AiModelConfig = {
  provider: "cloudflare",
  model: UNIFIED_MODEL_ID,
  apiToken: "ignored-in-gateway-mode",
};

// A Workers AI binding is a transport for the gateway's own routes, never for the account-level
// REST endpoint the unified catalog is served from.
const BINDING = { gateway: () => ({}) } as unknown as Ai;

function gatewayEnv(overrides: Partial<Cloudflare.Env> = {}): Cloudflare.Env {
  return {
    CF_AI_GATEWAY: "platform-gateway",
    CF_AI_GATEWAY_ACCOUNT_ID: "gateway-account-id",
    CF_AI_GATEWAY_API_TOKEN: "gateway-token",
    CF_AI_GATEWAY_PROVIDERS: "cloudflare",
    ...overrides,
  } as Cloudflare.Env;
}

type CapturedRequest = { url: string; headers: Headers };

const capturedRequests: CapturedRequest[] = [];

const fetchStub = (async (input: RequestInfo | URL, init?: RequestInit) => {
  const request = new Request(input as RequestInfo, init);
  capturedRequests.push({ url: request.url, headers: request.headers });
  // A non-retryable client error: pi turns it into an error-stop assistant message, and the
  // request stays captured for assertions.
  return Response.json({ error: { type: "bad_request", message: "stubbed" } }, { status: 400 });
}) as typeof fetch;

async function captureRequest(handle: ModelHandle): Promise<CapturedRequest> {
  const stream = await handle.stream(handle.model, {
    messages: [{ role: "user", content: "hello", timestamp: 0 }],
  }, { fetch: fetchStub, maxRetries: 0 });
  const message = await stream.result();
  expect(message.stopReason).toBe("error");
  expect(capturedRequests.length).toBeGreaterThan(0);
  return capturedRequests[0];
}

describe("suggested Cloudflare catalog", () => {
  // The window is the vendor's published 1M held back by 8%, and agent-compaction.ts computes its
  // trigger from this number -- so it is published state, not a tunable.
  it("offers exactly one Cloudflare model, the unified-catalog Qwen with its published window",
      () => {
    expect(SUGGESTED_MODELS.cloudflare).toEqual({
      [UNIFIED_MODEL_ID]: {
        name: "Qwen 3.7 Plus", contextWindow: 917504, outputLimit: WORKERS_AI_OUTPUT_LIMIT,
      },
    });
  });
});

describe("getModel unified-catalog routing", () => {
  beforeEach(() => {
    capturedRequests.length = 0;
  });

  // The binding satisfies AI Gateway mode's transport requirement, so nothing earlier objects --
  // which is exactly why this has to be caught here rather than per turn.
  it("refuses a unified-catalog model when the deployment has only the Workers AI binding", () => {
    expect(() => getModel(
        gatewayEnv({ CF_AI_GATEWAY_API_TOKEN: undefined, WORKERS_AI: BINDING }),
        UNIFIED_CONFIG, INITIATOR,
    )).toThrow("Unified-catalog models are served over the AI REST endpoint, which the Workers " +
        "AI binding cannot carry: set CF_AI_GATEWAY_API_TOKEN.");
  });

  it("routes a unified-catalog model to the account AI REST endpoint over OpenAI completions",
      () => {
    const handle = getModel(gatewayEnv(), UNIFIED_CONFIG, INITIATOR);

    expect(handle.model.id).toBe(UNIFIED_MODEL_ID);
    expect(handle.model.api).toBe("openai-completions");
    expect(handle.model.provider).toBe("cloudflare-workers-ai");
    expect(handle.model.baseUrl).toBe(
        "https://api.cloudflare.com/client/v4/accounts/gateway-account-id/ai/v1");
    // Published window and output limit, from the suggested catalog rather than pi's.
    expect(handle.model.contextWindow).toBe(917504);
    expect(handle.model.maxTokens).toBe(WORKERS_AI_OUTPUT_LIMIT);
    expect(handle.aiGatewayLogRoute).toEqual({
      gateway: "platform-gateway",
      accountId: "gateway-account-id",
      apiToken: "gateway-token",
    });
  });

  // The request leaves the gateway's data plane, so attribution rides a header instead of the URL.
  it("keeps logging and billing on the configured gateway via cf-aig-gateway-id", async () => {
    const handle = getModel(gatewayEnv(), UNIFIED_CONFIG, INITIATOR, {
      metadata: { source: "chat", gadgetId: "gadget-123", chatId: 7 },
    });

    const request = await captureRequest(handle);

    expect(request.url).toBe("https://api.cloudflare.com/client/v4/accounts/gateway-account-id/" +
        "ai/v1/chat/completions");
    expect(request.headers.get("cf-aig-gateway-id")).toBe("platform-gateway");
    expect(request.headers.get("authorization")).toBe("Bearer gateway-token");
    expect(request.headers.get("cf-aig-metadata")).toContain('"gadgetId":"gadget-123"');
  }, 15000);

  // Cost-log reads are same-account whenever the binding transport is active, even though this
  // model's inference cannot ride it.
  it("reads its cost logs over the binding when the deployment has one", () => {
    const handle = getModel(
        gatewayEnv({ WORKERS_AI: BINDING }), UNIFIED_CONFIG, INITIATOR);

    expect(handle.model.baseUrl).toBe(
        "https://api.cloudflare.com/client/v4/accounts/gateway-account-id/ai/v1");
    expect(handle.aiGatewayLogRoute).toEqual({ gateway: "platform-gateway" });
  });
});
