import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  getDevServerConfig,
  getWranglerPortFromBackendHost,
  seedGatekeeperVars,
} from "./dev-server-config.js";

describe("getWranglerPortFromBackendHost", () => {
  it("extracts a port from a localhost backend host", () => {
    assert.equal(getWranglerPortFromBackendHost("localhost:9000"), "9000");
  });

  it("extracts a port from an IPv6 backend host", () => {
    assert.equal(getWranglerPortFromBackendHost("[::1]:9001"), "9001");
  });

  it("returns null when the backend host has no port", () => {
    assert.equal(getWranglerPortFromBackendHost("localhost"), null);
  });

  it("rejects invalid ports", () => {
    assert.throws(
        () => getWranglerPortFromBackendHost("localhost:99999"),
        /VITE_BACKEND_HOST must include a valid port/);
  });

  it("rejects invalid IPv6 ports", () => {
    assert.throws(
        () => getWranglerPortFromBackendHost("[::1]:99999"),
        /VITE_BACKEND_HOST must include a valid port/);
  });

  it("rejects port zero", () => {
    assert.throws(
        () => getWranglerPortFromBackendHost("localhost:0"),
        /VITE_BACKEND_HOST must include a valid port/);
  });

  it("rejects invalid hosts", () => {
    assert.throws(
        () => getWranglerPortFromBackendHost("http://localhost:9000"),
        /VITE_BACKEND_HOST must include a valid host/);
  });
});

describe("getDevServerConfig", () => {
  it("uses VITE_BACKEND_HOST as the public host and Wrangler port", () => {
    assert.deepEqual(getDevServerConfig([], "localhost:9000"), {
      backendHost: "localhost:9000",
      wranglerPort: "9000",
    });
  });

  it("uses --port as the public host and Wrangler port", () => {
    assert.deepEqual(getDevServerConfig(["--port", "8899"]), {
      backendHost: "localhost:8899",
      wranglerPort: "8899",
    });
  });

  it("accepts --port=value", () => {
    assert.deepEqual(getDevServerConfig(["--port=8899"]), {
      backendHost: "localhost:8899",
      wranglerPort: "8899",
    });
  });

  for (const args of [["--port"], ["--port", "nope"], ["--port=0"], ["--port=65536"]]) {
    it(`rejects invalid arguments: ${args.join(" ")}`, () => {
      assert.throws(() => getDevServerConfig(args), /--port must be an integer between 1 and 65535/);
    });
  }
});

function seed(name, env, vars = {}) {
  const config = { vars };
  const warnings = seedGatekeeperVars(config, name, env);
  return { vars: config.vars, warnings };
}

describe("seedGatekeeperVars", () => {
  it("seeds CLIENT_ID/CLIENT_SECRET from the gatekeeper's shared shell vars", () => {
    assert.deepEqual(
        seed("gatekeeper-github", { GITHUB_CLIENT_ID: "id", GITHUB_CLIENT_SECRET: "secret" }),
        { vars: { CLIENT_ID: "id", CLIENT_SECRET: "secret" }, warnings: [] });
  });

  it("keeps credentials the gatekeeper's own config already defines", () => {
    assert.deepEqual(
        seed("gatekeeper-google",
            { GOOGLE_CLIENT_ID: "shell-id", GOOGLE_CLIENT_SECRET: "shell-secret" },
            { CLIENT_ID: "committed-id" }),
        { vars: { CLIENT_ID: "committed-id", CLIENT_SECRET: "shell-secret" }, warnings: [] });
  });

  it("seeds nothing when only one half of the credential pair is set", () => {
    assert.deepEqual(
        seed("gatekeeper-slack", { SLACK_CLIENT_ID: "id" }),
        { vars: {}, warnings: [] });
  });

  it("leaves a gatekeeper with no shared creds and no passthrough vars untouched", () => {
    assert.deepEqual(
        seed("gatekeeper-email", { GITHUB_CLIENT_ID: "id", GITHUB_CLIENT_SECRET: "secret" }),
        { vars: {}, warnings: [] });
  });

  it("lets the shell override a passthrough var the committed config already sets", () => {
    assert.deepEqual(
        seed("gatekeeper-mcp", { MCP_ALLOW_INSECURE: "true" }, { MCP_ALLOW_INSECURE: "false" }),
        { vars: { MCP_ALLOW_INSECURE: "true" }, warnings: [] });
  });

  it("seeds the Microsoft gatekeeper's credentials and renames the tenant id", () => {
    assert.deepEqual(
        seed("gatekeeper-microsoft", {
          MICROSOFT_CLIENT_ID: "id",
          MICROSOFT_CLIENT_SECRET: "secret",
          MICROSOFT_TENANT_ID: "tenant",
        }),
        {
          vars: { CLIENT_ID: "id", CLIENT_SECRET: "secret", TENANT_ID: "tenant" },
          warnings: [],
        });
  });

  it("seeds TENANT_ID even when the Microsoft client secret is missing", () => {
    assert.deepEqual(
        seed("gatekeeper-microsoft",
            { MICROSOFT_CLIENT_ID: "id", MICROSOFT_TENANT_ID: "tenant" }),
        { vars: { TENANT_ID: "tenant" }, warnings: [] });
  });

  for (const [label, tenantId] of [["missing", undefined], ["empty", ""], ["blank", "  "]]) {
    it(`warns and seeds no TENANT_ID when the tenant id is ${label} but the creds are set`, () => {
      const { vars, warnings } = seed("gatekeeper-microsoft", {
        MICROSOFT_CLIENT_ID: "id",
        MICROSOFT_CLIENT_SECRET: "secret",
        ...(tenantId === undefined ? {} : { MICROSOFT_TENANT_ID: tenantId }),
      });
      assert.deepEqual(vars, { CLIENT_ID: "id", CLIENT_SECRET: "secret" });
      assert.equal(warnings.length, 1);
      assert.match(warnings[0], /MICROSOFT_TENANT_ID is empty or not set/);
    });

    it(`stays quiet when the tenant id is ${label} and the gatekeeper is unconfigured`, () => {
      assert.deepEqual(
          seed("gatekeeper-microsoft", tenantId === undefined ? {} : { MICROSOFT_TENANT_ID: tenantId }),
          { vars: {}, warnings: [] });
    });
  }

  it("still passes an empty value through for a var that is not required", () => {
    assert.deepEqual(
        seed("gatekeeper-mcp", { MCP_ALLOW_INSECURE: "" }, { MCP_ALLOW_INSECURE: "false" }),
        { vars: { MCP_ALLOW_INSECURE: "" }, warnings: [] });
  });

  it("creates the vars object when the config has none", () => {
    const config = {};
    const warnings = seedGatekeeperVars(config, "gatekeeper-notion",
        { NOTION_CLIENT_ID: "id", NOTION_CLIENT_SECRET: "secret" });
    assert.deepEqual(config.vars, { CLIENT_ID: "id", CLIENT_SECRET: "secret" });
    assert.deepEqual(warnings, []);
  });
});
