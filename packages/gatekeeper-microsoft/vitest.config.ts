import { fileURLToPath } from "node:url";
import capnwebValidate from "capnweb-validate/vite";
import { defineConfig } from "vitest/config";

export default defineConfig({
  // The worker's RPC entrypoints carry `@validateRpc()`, which refuses to run untransformed.
  plugins: [capnwebValidate()],
  test: {
    include: ["__tests__/*.test.ts"],
    environment: "node",
    // Supplies the one workerd-only crypto primitive the OAuth nonce check calls.
    setupFiles: ["./__tests__/stubs/workerd-crypto-shim.ts"],
    alias: {
      // Lets the worker module, which declares a Durable Object and RPC entrypoints, be imported at
      // all. See the stub for what it does and does not provide.
      "cloudflare:workers": fileURLToPath(
        new URL("./__tests__/stubs/cloudflare-workers.ts", import.meta.url)),
    },
  },
});
