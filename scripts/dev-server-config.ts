/**
 * The port a `VITE_BACKEND_HOST` names, as a string, or null when it names no port. Throws on a
 * value that is not a bare `host[:port]`.
 */
export function getWranglerPortFromBackendHost(backendHost: string): string | null {
  const trimmed = backendHost.trim();
  if (!trimmed) return null;
  if (trimmed.includes("://")) {
    throw new Error("VITE_BACKEND_HOST must include a valid host with an optional port.");
  }

  let url: URL;
  try {
    url = new URL(`http://${trimmed}`);
  } catch {
    if (/(^.*\]:|^[^:]+:)[^:]+$/.test(trimmed)) {
      throw new Error("VITE_BACKEND_HOST must include a valid port between 1 and 65535.");
    }
    throw new Error("VITE_BACKEND_HOST must include a valid host with an optional port.");
  }

  if (!url.port) return null;

  const port = Number(url.port);
  if (port < 1) {
    throw new Error("VITE_BACKEND_HOST must include a valid port between 1 and 65535.");
  }

  return url.port;
}

/**
 * Maps a gatekeeper name to the shell env vars whose values seed its CLIENT_ID / CLIENT_SECRET.
 * A single OAuth app can then drive both the sign-in flow and the gatekeeper's own capabilities.
 */
export const SHARED_GATEKEEPER_CREDS: Record<string, { id: string; secret: string }> = {
  "gatekeeper-github": { id: "GITHUB_CLIENT_ID", secret: "GITHUB_CLIENT_SECRET" },
  "gatekeeper-google": { id: "GOOGLE_CLIENT_ID", secret: "GOOGLE_CLIENT_SECRET" },
  "gatekeeper-cloudflare": { id: "CLOUDFLARE_OAUTH_CLIENT_ID", secret: "CLOUDFLARE_OAUTH_CLIENT_SECRET" },
  "gatekeeper-supabase": { id: "SUPABASE_CLIENT_ID", secret: "SUPABASE_CLIENT_SECRET" },
  "gatekeeper-notion": { id: "NOTION_CLIENT_ID", secret: "NOTION_CLIENT_SECRET" },
  "gatekeeper-zoominfo": { id: "ZOOMINFO_CLIENT_ID", secret: "ZOOMINFO_CLIENT_SECRET" },
  "gatekeeper-confluence": { id: "CONFLUENCE_CLIENT_ID", secret: "CONFLUENCE_CLIENT_SECRET" },
  "gatekeeper-slack": { id: "SLACK_CLIENT_ID", secret: "SLACK_CLIENT_SECRET" },
  "gatekeeper-microsoft": { id: "MICROSOFT_CLIENT_ID", secret: "MICROSOFT_CLIENT_SECRET" },
};

/** A passthrough entry whose worker-side name differs from the shell var it reads. */
export interface PassthroughVar {
  /** Shell (or `.dev.vars`) variable the value comes from. */
  from: string;
  /** Variable name the worker reads. */
  to: string;
  /** Warn when the gatekeeper has its shared credentials but this var is blank or unset. */
  required?: boolean;
}

/**
 * Deployment-configured vars a gatekeeper reads that its committed `wrangler.jsonc` deliberately
 * leaves unset, passed through from the shell or the root `.dev.vars`.
 *
 * Without this the only way to point the portal connector somewhere for local testing is to edit a
 * tracked file, and a URL committed there becomes the default for everyone who deploys this repo.
 * `.dev.vars` is gitignored, so it cannot leave the machine. Secrets travel the same way
 * `CLIENT_SECRET` already does, via SHARED_GATEKEEPER_CREDS above.
 *
 * An entry is either a var name used verbatim, or `{ from, to }` when the worker's var name
 * differs from the shell var name. `required: true` asks for a warning when the gatekeeper has its
 * shared credentials but not this var: the developer is clearly setting that gatekeeper up, and the
 * var's absence otherwise surfaces only as a "not configured" page that names nothing.
 */
export const PASSTHROUGH_GATEKEEPER_VARS: Record<string, (string | PassthroughVar)[]> = {
  "gatekeeper-mcp-portal": [
    "MCP_PORTAL_URL", "MCP_PORTAL_NAME", "MCP_PORTAL_AUTH", "MCP_PORTAL_TOKEN",
    "MCP_PORTAL_TRUST_ANNOTATIONS", "MCP_ALLOW_INSECURE",
  ],
  "gatekeeper-mcp": ["MCP_ALLOW_INSECURE"],
  // Entra pins every OAuth endpoint to one tenant, so the Microsoft gatekeeper cannot build an
  // authorize URL without it — it is as mandatory as the client credentials.
  "gatekeeper-microsoft": [{ from: "MICROSOFT_TENANT_ID", to: "TENANT_ID", required: true }],
};

/**
 * Seeds one gatekeeper's dev-only vars into `config.vars` (mutated in place) and returns the
 * warnings the caller should print. Credentials already present in the gatekeeper's own config
 * win; passthrough vars are shell-wins, so `MCP_ALLOW_INSECURE=true` overrides the `"false"` in a
 * committed wrangler.jsonc without editing a tracked file.
 */
export function seedGatekeeperVars(
  config: { vars?: Record<string, string | undefined> },
  gatekeeperName: string,
  env: Record<string, string | undefined>,
): string[] {
  const warnings: string[] = [];
  const vars = (config.vars ??= {});

  const shared = SHARED_GATEKEEPER_CREDS[gatekeeperName];
  const credsSeeded = Boolean(shared && env[shared.id] && env[shared.secret]);
  if (credsSeeded) {
    if (vars.CLIENT_ID === undefined) vars.CLIENT_ID = env[shared.id];
    if (vars.CLIENT_SECRET === undefined) vars.CLIENT_SECRET = env[shared.secret];
  }

  // Deliberately outside the credential guard: a non-credential var must still reach the worker
  // when only part of an OAuth app is configured. Folding it into the guard would let a missing
  // client secret silently take the rest of the gatekeeper's configuration with it.
  for (const entry of PASSTHROUGH_GATEKEEPER_VARS[gatekeeperName] ?? []) {
    const { from, to, required }: PassthroughVar =
        typeof entry === "string" ? { from: entry, to: entry } : entry;
    const value = env[from];
    // A `required` var counts as absent when it is blank: a worker that rejects an empty TENANT_ID
    // the same way it rejects a missing one gains nothing from seeding `""`, and seeding it would
    // suppress the warning — leaving the same unexplained "not configured" page this warning
    // exists to name. An empty line in `.dev.vars` is the usual way to arrive here. Optional
    // entries keep passing empty values through, since blank can be a meaningful override.
    const present = required ? Boolean(value?.trim()) : value !== undefined;
    if (present) {
      vars[to] = value;
    } else if (required && credsSeeded) {
      warnings.push(
          `${gatekeeperName}: ${from} is empty or not set, so the worker has no ${to} and every ` +
          `request to it will answer "not configured".`);
    }
  }

  return warnings;
}

/**
 * Resolve where the dev server's backend lives: an explicit `--port` wins, else
 * `VITE_BACKEND_HOST`, else `localhost:8787`.
 */
export function getDevServerConfig(args: readonly string[], envBackendHost?: string): {
  backendHost: string;
  wranglerPort: string | null;
} {
  let commandLinePort: string | null = null;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg !== "--port" && !arg.startsWith("--port=")) continue;

    if (commandLinePort !== null) {
      throw new Error("--port may only be specified once.");
    }

    const value = arg === "--port" ? args[++i] : arg.slice("--port=".length);
    if (!/^\d+$/.test(value ?? "") || Number(value) < 1 || Number(value) > 65535) {
      throw new Error("--port must be an integer between 1 and 65535.");
    }
    commandLinePort = String(Number(value));
  }

  if (commandLinePort !== null) {
    return {
      backendHost: `localhost:${commandLinePort}`,
      wranglerPort: commandLinePort,
    };
  }

  const backendHost = envBackendHost?.trim() || "localhost:8787";
  return {
    backendHost,
    wranglerPort: getWranglerPortFromBackendHost(backendHost),
  };
}
