import { WorkerEntrypoint, DurableObject, RpcStub, RpcTarget } from "cloudflare:workers";
import { skipRpcValidation, validateRpc } from "capnweb-validate";
import {
  AccountDescription, Gatekeeper, GatekeeperConnectCallback, GatekeeperConnectOptions,
  GatekeeperUser, GatekeeperUserVerifier, GatekeeperVendor as GatekeeperVendorIface,
  ResourceConfiguratorFrame, SupportedResource, VendorDescription, stripTrailingSlashes,
} from "@gadgets/workshop-shared/gatekeeper";
import {
  AccessTokenRequest, EntraAccessToken, EntraTokenGrant, IdTokenClaims, ProfilePhoto, TokenFailure,
  authorizeEndpoint, exchangeAuthCode, fetchAvatarDataUrl, fetchGraphProfile, fetchProfilePhoto,
  fetchVerifiedDomains, grantCoversScopes, refreshAccessToken, resolveVerifiedEmail,
} from "./microsoft-api";
import type { OutlookMailGatekeeperImplProps } from "./outlook-mail";
import type {
  OutlookMailConfiguratorRpc,
} from "./configurator/outlook-mail-configurator-types";
import MICROSOFT_LOGO_SVG from "./microsoft-logo.svg";
import OUTLOOK_MAIL_CONFIGURATOR_HTML from "./generated/outlook-mail-configurator-ui.txt";
import TYPES_CODE from "./types.txt";
import { obsContext } from "./observability.js";

export { OutlookMailGatekeeperImpl } from "./outlook-mail";

// Vendor id = GATEKEEPER_<NAME> binding suffix (lowercased).
const VENDOR_ID = "microsoft";
const logger = obsContext.createLogger({
  component: "gatekeeper.microsoft", vendorId: VENDOR_ID,
});

// A nonce stored in UserAccount KV to protect the OAuth flow. Only one nonce is active at a time;
// the `stage` field tracks where we are in the flow.
type StoredNonce = {
  value: string;
  expiresAt: number;
  stage: "initiation" | "oauth";
};

// The last mint that failed permanently, kept in storage so an evicted object doesn't rediscover a
// dead account by hammering the token endpoint.
type StoredMintFailure = {
  message: string;
  at: number;
};

type CachedVerifiedDomains = {
  domains: string[];
  fetchedAt: number;
};

const NONCE_BYTES = 32;
const INITIATION_NONCE_LIFETIME_MS = 10 * 60 * 1000;  // 10 minutes
const OAUTH_NONCE_LIFETIME_MS = 10 * 60 * 1000;    // 10 minutes

// How long an unfinished flow may sit before the account deletes itself.
const ABANDONED_FLOW_LIFETIME_MS = 60 * 60 * 1000;

// How long a completed sign-in-only grant survives. The caller reads the email during complete(),
// so this only needs to outlast that call.
const AUTH_GRANT_LIFETIME_MS = 2 * 60 * 1000;

// Ceilings on the round trips that run while holding the credential mutex. Each must be bounded: an
// unbounded hang keeps the mutex, and every caller waiting for a token then queues behind it.
const TOKEN_MINT_TIMEOUT_MS = 20 * 1000;
const AUTH_CODE_EXCHANGE_TIMEOUT_MS = 30 * 1000;
const GRAPH_TIMEOUT_MS = 20 * 1000;

// How long a permanent mint failure suppresses further attempts. Long enough to absorb the burst a
// revocation produces (every outstanding token fails at once, so callers arrive within
// milliseconds), short enough that the account recovers on its own once an administrator fixes the
// cause — a restored consent leaves nothing for us to observe, so we have to re-ask Entra
// eventually.
const MINT_FAILURE_COOLDOWN_MS = 60 * 1000;

// The tenant's verified domains change only when an administrator adds or removes a domain, so a
// day-old answer is fine. This spares repeat lookups on a persistent account; it does NOT speed up
// sign-in, since each login runs on a freshly minted account object that self-destructs minutes
// later, so every login pays for its own lookup.
const VERIFIED_DOMAIN_CACHE_MS = 24 * 60 * 60 * 1000;

// How far ahead of its recorded expiry an access token is treated as already expired, so a token
// doesn't lapse mid-request.
const ACCESS_TOKEN_EXPIRY_SAFETY_MS = 60 * 1000;

/** Longest attacker-influenced value echoed back on the OAuth error page. */
const MAX_OAUTH_ERROR_CHARS = 500;

function hexEncode(bytes: Uint8Array): string {
  return [...bytes].map(b => b.toString(16).padStart(2, "0")).join("");
}

function generateNonce(): string {
  return hexEncode(crypto.getRandomValues(new Uint8Array(NONCE_BYTES)));
}

function constantTimeEqual(a: string, b: string): boolean {
  let encoder = new TextEncoder();
  let bufA = encoder.encode(a);
  let bufB = encoder.encode(b);
  if (bufA.byteLength !== bufB.byteLength) return false;
  return crypto.subtle.timingSafeEqual(bufA, bufB);
}

/**
 * Declare optional environment variables here since they may be omitted from wrangler.jsonc.
 */
export type Env = Cloudflare.Env & {
  /**
   * Base URL (protocol+host+optional path) at which the default fetch handler is served. Should
   * NOT include a trailing slash. Omit for localhost dev server.
   */
  BASE_URL?: string;
  /**
   * Entra ID app registration credentials (wrangler secrets / .dev.vars); not in wrangler.jsonc.
   */
  CLIENT_ID?: string;
  CLIENT_SECRET?: string;
  /**
   * The single Entra tenant this deployment signs users in from. Pins both the OAuth endpoints and
   * the id_token `tid` check.
   */
  TENANT_ID?: string;
}

function getBaseUrl(env: Env) {
  return stripTrailingSlashes(env.BASE_URL || "http://localhost:8787/gatekeeper/microsoft");
}

function getBasePath(env: Env) {
  const path = new URL(getBaseUrl(env)).pathname;
  return path === "/" ? "" : path;
}

type OAuthConfig = {
  clientId: string;
  clientSecret: string;
  tenantId: string;
};

function getOAuthConfig(env: Env): OAuthConfig | null {
  if (!env.CLIENT_ID || !env.CLIENT_SECRET || !env.TENANT_ID) return null;
  return { clientId: env.CLIENT_ID, clientSecret: env.CLIENT_SECRET, tenantId: env.TENANT_ID };
}

// =======================================================================================

const SELF_CLOSING_HTML = `<!DOCTYPE html>
<html lang="en">
  <body>
    <script type="text/javascript">window.close();</script>
    <p>Authorization complete. You may close this tab and return to Cloudflare OS.
  </body>
</html>`;

const INVALID_LINK_HTML = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Authorization Link Expired</title>
  </head>
  <body style="font-family: system-ui, -apple-system, sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; background: #f5f5f5;">
    <div style="max-width: 520px; padding: 2rem; background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); text-align: center;">
      <h1 style="color: #d97706; font-size: 1.5rem; margin: 0 0 1rem 0;">Authorization Link Expired</h1>
      <p style="color: #555; line-height: 1.6; margin: 0 0 1.5rem 0;">This authorization link is invalid or has expired. Please return to Cloudflare OS and try again.</p>
      <button onclick="window.close()" style="padding: 0.5rem 1.5rem; background: #d97706; color: white; border: none; border-radius: 4px; font-size: 1rem; cursor: pointer;">Close</button>
    </div>
  </body>
</html>`;

const NOT_CONFIGURED_HTML = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Configuration Required</title>
  </head>
  <body style="font-family: system-ui, -apple-system, sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; background: #f5f5f5;">
    <div style="max-width: 520px; padding: 2rem; background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); text-align: center;">
      <h1 style="color: #d97706; font-size: 1.5rem; margin: 0 0 1rem 0;">Microsoft Gatekeeper Not Configured</h1>
      <p style="color: #555; line-height: 1.6; margin: 0;">Please see the README.md for instructions on configuring an Entra ID application (client ID, client secret, and tenant ID) so that this Cloudflare OS instance can access Microsoft 365.</p>
    </div>
  </body>
</html>`;

// Scopes requested for a sign-in-only connection. `offline_access` is deliberately absent: the
// grant does one Graph read and then self-destructs, so a refresh token would be blast radius with
// no purpose. Entra issues an access token from the code exchange regardless; `offline_access` only
// gates the refresh token.
const AUTH_SCOPES = ["openid", "profile", "email", "User.Read"];

// Scopes requested for a persistent connection, which has to survive past one access token.
const IDENTITY_SCOPES = [...AUTH_SCOPES, "offline_access"];

const OUTLOOK_MAIL_RESOURCE: SupportedResource = {
  urlPattern: "https://outlook.office.com/mail/*",
  title: "Outlook Mailbox",
  description: "Read emails, organize folders, and draft replies.",
  grantable: true,
};

// `Mail.ReadWrite` covers reading messages and folders, flipping read state, moving messages, and
// creating drafts. `Mail.Send` is deliberately absent: this gatekeeper never sends mail, it leaves
// drafts for the user.
const RESOURCE_SCOPES: {resource: SupportedResource, scopes: string[]}[] = [
  { resource: OUTLOOK_MAIL_RESOURCE, scopes: ["Mail.ReadWrite"] },
];

const SUPPORTED_RESOURCES: SupportedResource[] = RESOURCE_SCOPES.map(entry => entry.resource);

function validateResourceUrlPatterns(resourceUrlPatterns?: string[]): void {
  if (resourceUrlPatterns === undefined) return;

  let knownPatterns = new Set(RESOURCE_SCOPES.map(entry => entry.resource.urlPattern));
  let unknownPatterns = resourceUrlPatterns.filter(pattern => !knownPatterns.has(pattern));
  if (unknownPatterns.length > 0) {
    throw new Error(`Unknown grantable resource URL pattern(s): ${unknownPatterns.join(", ")}`);
  }
}

// The OAuth scopes to request for the given grantable resource `urlPattern`s. Undefined means
// "every resource this vendor offers".
function resourceUrlPatternsToOAuthScopes(resourceUrlPatterns?: string[]): string[] {
  validateResourceUrlPatterns(resourceUrlPatterns);

  let scopes = new Set<string>(IDENTITY_SCOPES);
  for (let entry of RESOURCE_SCOPES) {
    if (resourceUrlPatterns === undefined ||
        resourceUrlPatterns.includes(entry.resource.urlPattern)) {
      for (let scope of entry.scopes) scopes.add(scope);
    }
  }
  return [...scopes];
}

// The resources covered by a grant that reported `grantedOAuthScopes`. Comparison is normalized —
// see `grantCoversScopes`, which is what keeps a resource-qualified or differently-cased response
// from looking like a missing permission and re-prompting for consent forever.
function grantedResourcesFromScopes(grantedOAuthScopes: string[]): string[] {
  return RESOURCE_SCOPES
      .filter(entry => grantCoversScopes(grantedOAuthScopes, entry.scopes))
      .map(entry => entry.resource.urlPattern);
}

const MICROSOFT_LOGO_URL = `data:image/svg+xml,${encodeURIComponent(MICROSOFT_LOGO_SVG)}`;

function truncateOAuthError(value: string): string {
  return value.length > MAX_OAUTH_ERROR_CHARS
      ? `${value.slice(0, MAX_OAUTH_ERROR_CHARS)}…`
      : value;
}

/**
 * Main HTTP UI entrypoint. We only use this to initiate and complete OAuth requests to Entra ID.
 */
export default {
  async fetch(req: Request, env: Env, ctx: ExecutionContext) {
    let url = new URL(req.url);
    let basePath = getBasePath(env);
    if (!url.pathname.startsWith(basePath + "/") && url.pathname !== basePath) {
      throw new Error(`Request path ${url.pathname} does not match BASE_URL path ${basePath}`);
    }
    let relPath = url.pathname.slice(basePath.length);
    let path = relPath.slice(1).split("/");

    if (path.length === 2 && path[0].length === 64 && path[1].length === NONCE_BYTES * 2) {
      let config = getOAuthConfig(env);
      if (!config) {
        return new Response(NOT_CONFIGURED_HTML, {
          headers: {
            "Content-Type": "text/html; charset=utf-8"
          }
        });
      }

      let doId = path[0];
      let initiationNonce = path[1];
      let stub = ctx.exports.UserAccount.get(ctx.exports.UserAccount.idFromString(doId));
      let begun = await stub.beginOAuthFlow(initiationNonce);
      if (begun === null) {
        return new Response(INVALID_LINK_HTML, {
          headers: { "Content-Type": "text/html; charset=utf-8" }
        });
      }

      let newUrl = new URL(authorizeEndpoint(config.tenantId));
      newUrl.searchParams.set("client_id", config.clientId);
      newUrl.searchParams.set("redirect_uri", getBaseUrl(env) + "/oauth");
      // Code only. Asking for an id_token here would deliver one over the front channel, where it
      // is attacker-supplied data; the only id_token we trust is the one the token endpoint hands
      // back on the exchange below.
      newUrl.searchParams.set("response_type", "code");
      newUrl.searchParams.set("response_mode", "query");
      newUrl.searchParams.set("scope", begun.scopes.join(" "));
      newUrl.searchParams.set("state", `${doId}:${begun.oauthNonce}`);

      return Response.redirect(newUrl.toString(), 302);
    } else if (relPath === "/oauth") {
      // Completion redirect.

      let error = url.searchParams.get("error");
      if (error) {
        // text/plain, explicitly: `error` and `error_description` are query data an attacker can
        // choose, and this page is served from the deployment's own origin, so rendering them as
        // HTML would be reflected XSS. Both are length-capped for the same reason.
        let description = url.searchParams.get("error_description") ?? "";
        return new Response(
            `${truncateOAuthError(error)}: ${truncateOAuthError(description)}`, {
              status: 400,
              headers: {
                "Content-Type": "text/plain; charset=utf-8",
                "X-Content-Type-Options": "nosniff",
              },
            });
      }

      let state = url.searchParams.get("state");
      if (!state) return new Response("Error: no 'state' provided");
      let colonIdx = state.indexOf(":");
      if (colonIdx < 0) return new Response("Error: malformed state");
      let doId = state.slice(0, colonIdx);
      let oauthNonce = state.slice(colonIdx + 1);

      let code = url.searchParams.get("code");
      if (!code) return new Response("Error: no 'code' provided");

      let userObjectId = ctx.exports.UserAccount.idFromString(doId);
      let stub: DurableObjectStub<UserAccount> = ctx.exports.UserAccount.get(userObjectId);
      if (!await stub.acceptAuthCode(code, oauthNonce)) {
        return new Response(INVALID_LINK_HTML, {
          headers: { "Content-Type": "text/html; charset=utf-8" }
        });
      }
      return new Response(SELF_CLOSING_HTML, {
        headers: {
          "Content-Type": "text/html; charset=utf-8"
        }
      });
    } else {
      return new Response("Not Found", {status: 404});
    }
  }
}

// =======================================================================================

// Top-level API exposed to the Workshop.
@validateRpc()
export class GatekeeperVendor extends WorkerEntrypoint<Env> implements GatekeeperVendorIface {
  status() {
    return "Microsoft Gatekeeper";
  }

  async describe(): Promise<VendorDescription> {
    return {
      displayName: "Microsoft",
      url: "https://www.microsoft.com",
      logo: { url: MICROSOFT_LOGO_URL },
      color: "#eef1f6",
      tagline: "Sign in with your Microsoft 365 work account",
      description:
          "Connect your Microsoft 365 account so Cloudflare OS can confirm who you are using your " +
          "organization's Entra ID directory. Sign-in uses the account you already have — no " +
          "separate password — and the connection is limited to reading your own profile.",
      providesAuth: true,
      providesAuthProfile: true,
    };
  }

  async connectAccount(callback: Fetcher<GatekeeperConnectCallback>,
                       options?: GatekeeperConnectOptions): Promise<{url: string}> {
    validateResourceUrlPatterns(options?.resourceUrlPatterns);

    let userObjectId = this.ctx.exports.UserAccount.newUniqueId();
    let initiationNonce = generateNonce();

    let authOnly = options?.scopes === "auth";
    let requestedScopes = authOnly
        ? AUTH_SCOPES
        : resourceUrlPatternsToOAuthScopes(options?.resourceUrlPatterns);
    await this.ctx.exports.UserAccount.get(userObjectId)
        .setCallback(callback, initiationNonce, requestedScopes, authOnly);

    return {
      url: `${getBaseUrl(this.env)}/${userObjectId.toString()}/${initiationNonce}`
    };
  }

  async newUser(): Promise<Fetcher<GatekeeperUser>> {
    let userObjectId = this.ctx.exports.UserAccount.newUniqueId();
    let props: GatekeeperUserImplProps = { userObjectId: userObjectId.toString() };
    return this.ctx.exports.GatekeeperUserImpl({props});
  }

  async getSupportedResources(): Promise<SupportedResource[]> {
    return SUPPORTED_RESOURCES;
  }

  async getTypeScriptTypes(): Promise<string> {
    return TYPES_CODE;
  }
}

export class UserAccount extends DurableObject<Env> {
  // Serialize minting, reconnect, and revoke against each other. Minting is a network round trip, so
  // without this a single invalidated token has every concurrent caller mint its own — a burst
  // against Entra's token endpoint that may get throttled, turning a recoverable rejection into a
  // hard failure. It also keeps a mint from interleaving with credentials being replaced or wiped,
  // which matters more here than for providers that never rotate: the mint both reads and replaces
  // the stored refresh token.
  //
  // A promise chain rather than blockConcurrencyWhile: that would freeze the whole object for the
  // duration of the fetch, and an exception or a 30s overrun inside it resets the Durable Object.
  #credentialUpdate: Promise<void> = Promise.resolve();

  async #updateCredentials<T>(operation: () => Promise<T>): Promise<T> {
    let previous = this.#credentialUpdate;
    let release!: () => void;
    this.#credentialUpdate = new Promise(resolve => { release = resolve; });
    await previous;
    try {
      return await operation();
    } finally {
      release();
    }
  }

  #requireConfig(): OAuthConfig {
    let config = getOAuthConfig(this.env);
    if (!config) {
      throw new Error("The Microsoft Gatekeeper is not configured.");
    }
    return config;
  }

  // Persist a freshly minted grant. Storing the returned refresh token is mandatory, not an
  // optimization: Entra retires the presented token on every mint, so keeping the old value strands
  // the account the moment the retired token stops being honored.
  #storeGrant(grant: EntraTokenGrant): void {
    if (grant.refreshToken) this.ctx.storage.kv.put<string>("refreshToken", grant.refreshToken);
    this.ctx.storage.kv.put<EntraAccessToken>("accessToken", grant.accessToken);
    if (grant.grantedScopes.length > 0) {
      this.ctx.storage.kv.put<string[]>("grantedScopes", grant.grantedScopes);
    }
    if (grant.idTokenClaims) {
      this.ctx.storage.kv.put<IdTokenClaims>("idTokenClaims", grant.idTokenClaims);
    }
  }

  // Drop every piece of a grant. Callers must hold the credential mutex.
  #clearGrant(): void {
    this.ctx.storage.kv.delete("refreshToken");
    this.ctx.storage.kv.delete("accessToken");
    this.ctx.storage.kv.delete("idTokenClaims");
    this.ctx.storage.kv.delete("grantedScopes");
    this.ctx.storage.kv.delete("grantScopes");
  }

  async setCallback(
      callback: Fetcher<GatekeeperConnectCallback>, initiationNonce: string,
      requestedScopes: string[], authOnly?: boolean) {
    // If the flow never completes, delete this object.
    if (!this.ctx.storage.kv.get<string>("refreshToken")) {
      this.ctx.storage.setAlarm(Date.now() + ABANDONED_FLOW_LIFETIME_MS);
    }

    this.ctx.storage.kv.put("callback", callback);
    this.ctx.storage.kv.put<string[]>("requestedScopes", requestedScopes);
    // Sign-in-only grants are transient: dropped shortly after the email is read.
    this.ctx.storage.kv.put<boolean>("authOnly", authOnly ?? false);
    this.ctx.storage.kv.put<StoredNonce>("nonce", {
      value: initiationNonce,
      expiresAt: Date.now() + INITIATION_NONCE_LIFETIME_MS,
      stage: "initiation",
    });
  }

  /**
   * Prepare this account for a reconnect flow. The next acceptAuthCode() call will replace the
   * existing refresh token and notify via credentialsRestored() instead of complete().
   *
   * `requestedScopes` is the full set of OAuth scopes to request on the reauthorization.
   */
  async prepareReconnect(initiationNonce: string, requestedScopes: string[]) {
    this.ctx.storage.kv.put<boolean>("reconnecting", true);
    this.ctx.storage.kv.put<string[]>("requestedScopes", requestedScopes);
    this.ctx.storage.kv.put<StoredNonce>("nonce", {
      value: initiationNonce,
      expiresAt: Date.now() + INITIATION_NONCE_LIFETIME_MS,
      stage: "initiation",
    });
  }

  /**
   * The scopes this account consented to, which a reconnect re-requests so it doesn't narrow
   * access.
   */
  async getGrantScopes(): Promise<string[]> {
    return this.ctx.storage.kv.get<string[]>("grantScopes") ?? IDENTITY_SCOPES;
  }

  /**
   * The grantable resource `urlPattern`s this account's consent actually covers, used to decide
   * whether ensureResources() has to expand the grant. Derived from the scopes Entra reported, not
   * from the scopes we asked for: a user can decline part of a consent screen.
   */
  async getGrantedResourceUrlPatterns(): Promise<string[]> {
    return grantedResourcesFromScopes(this.ctx.storage.kv.get<string[]>("grantedScopes") ?? []);
  }

  /**
   * The id_token claims recorded at grant time. Only ever populated from a token-endpoint response.
   */
  async getIdTokenClaims(): Promise<IdTokenClaims | undefined> {
    return this.ctx.storage.kv.get<IdTokenClaims>("idTokenClaims");
  }

  /**
   * Called by the fetch handler when the user visits the initiation URL. Verifies the initiation
   * nonce, consumes it, and returns a fresh OAuth nonce plus the scopes to request. Returns null if
   * the nonce is invalid or expired.
   */
  async beginOAuthFlow(initiationNonce: string): Promise<{oauthNonce: string, scopes: string[]} | null> {
    let stored = this.ctx.storage.kv.get<StoredNonce>("nonce");
    if (!stored || stored.stage !== "initiation" ||
        Date.now() >= stored.expiresAt || !constantTimeEqual(stored.value, initiationNonce)) {
      return null;
    }

    // Replace the consumed initiation nonce with a fresh OAuth nonce.
    let oauthNonce = generateNonce();
    this.ctx.storage.kv.put<StoredNonce>("nonce", {
      value: oauthNonce,
      expiresAt: Date.now() + OAUTH_NONCE_LIFETIME_MS,
      stage: "oauth",
    });
    let scopes = this.ctx.storage.kv.get<string[]>("requestedScopes") ?? IDENTITY_SCOPES;
    return {oauthNonce, scopes};
  }

  /**
   * Returns false if the OAuth nonce is invalid or expired.
   */
  async acceptAuthCode(code: string, oauthNonce: string): Promise<boolean> {
    // Verify and consume the OAuth nonce.
    let stored = this.ctx.storage.kv.get<StoredNonce>("nonce");
    if (!stored || stored.stage !== "oauth" ||
        Date.now() >= stored.expiresAt || !constantTimeEqual(stored.value, oauthNonce)) {
      return false;
    }
    this.ctx.storage.kv.delete("nonce");

    let config = this.#requireConfig();

    // The credential swap is serialized against minting and revoke, but the callbacks below are
    // not: they are outbound RPCs that can re-enter this object, and awaiting one while holding the
    // mutex would deadlock. So the locked section returns what the notifications need and the
    // notifications happen after it releases.
    let completion = await this.#updateCredentials(async () => {
      let callback = this.ctx.storage.kv.get<Fetcher<GatekeeperConnectCallback>>("callback");
      if (!callback) {
        // Must have timed out.
        throw new Error("Took too long to complete the authorization. Please try again.");
      }

      let scopes = this.ctx.storage.kv.get<string[]>("requestedScopes") ?? IDENTITY_SCOPES;
      let authOnly = this.ctx.storage.kv.get<boolean>("authOnly") ?? false;
      let reconnecting = !!this.ctx.storage.kv.get<boolean>("reconnecting");

      let grant = await exchangeAuthCode({
        ...config,
        code,
        redirectUri: getBaseUrl(this.env) + "/oauth",
        scopes,
      }, AbortSignal.timeout(AUTH_CODE_EXCHANGE_TIMEOUT_MS));

      // A sign-in-only grant is expected to come back without one (it never asked for
      // `offline_access`); a persistent connection without one would die with its first access
      // token.
      if (!authOnly && !grant.refreshToken) {
        throw new Error("Microsoft did not return a refresh token for this connection.");
      }

      // Re-authorizing must return the SAME directory principal. Nothing stops a user from picking
      // a different Microsoft account at the consent screen, and adopting it would silently repoint
      // an established connection — the Workshop's records (and any resources bound through this
      // account) still name the original identity, and a mailbox bound through it would start
      // serving someone else's mail. `oid` is the tenant-immutable handle for that principal:
      // unlike the address, it survives renames, so a mismatch is a real identity swap.
      //
      // Fail closed on a missing value on either side. Entra always returns `oid` for an
      // openid-scoped grant, so "no principal recorded" is not a legitimate established account —
      // it is a state we cannot verify, and an unverifiable reconnect is exactly what this check
      // exists to refuse. The grant is left untouched, as if the user had refused consent.
      if (reconnecting) {
        let priorObjectId = this.ctx.storage.kv.get<IdTokenClaims>("idTokenClaims")?.oid;
        if (!priorObjectId || grant.idTokenClaims?.oid !== priorObjectId) {
          throw new Error(
              "This connection belongs to a different Microsoft account. Sign in with the " +
              "account that was originally connected.");
        }
      }

      this.#storeGrant(grant);
      this.ctx.storage.kv.put<string[]>("grantScopes", scopes);
      this.ctx.storage.kv.delete("requestedScopes");
      // These credentials are new, so any recorded permanent failure no longer applies.
      this.ctx.storage.kv.delete("mintFailure");

      if (reconnecting) this.ctx.storage.kv.delete("reconnecting");
      return { callback, reconnecting, authOnly };
    });

    let callback = completion.callback;
    if (completion.reconnecting) {
      // Reconnect flow: credentials replaced above, notify restoration.
      await callback.credentialsRestored();
    } else {
      // Initial connect flow: create the user entrypoint and notify completion.
      //
      // No expiry is reported. The contract's `expiresAt` means "when the credentials stop being
      // refreshable", and Entra publishes no lifetime for a refresh token — it is invalidated by
      // inactivity, policy, or revocation, none of which are visible in advance. Credential death
      // is reported when it happens, via credentialsExpired().
      try {
        let props: GatekeeperUserImplProps = { userObjectId: this.ctx.id.toString() };
        await callback.complete(this.ctx.exports.GatekeeperUserImpl({props}));
      } catch (err) {
        // Nobody received a usable account, so the whole grant is orphaned — not just the refresh
        // token. Leaving the access token behind would let a concurrent mint publish it for an
        // account that never completed. Serialized like every other credential mutation.
        await this.#updateCredentials(async () => { this.#clearGrant(); });
        throw err;
      }
      // Sign-in-only grants are transient: the caller has read the email via complete(), so
      // schedule a prompt self-destruct.
      if (completion.authOnly) {
        this.ctx.storage.setAlarm(Date.now() + AUTH_GRANT_LIFETIME_MS);
      }
    }

    return true;
  }

  hasRefreshToken() {
    return this.ctx.storage.kv.get<string>("refreshToken") !== undefined;
  }

  /**
   * Whether the stored token still satisfies this request, i.e. can be served without minting.
   *
   * A `staleToken` request comes from a caller that just had a 401. It must not be answered from the
   * expiry check — the whole point is that the token was rejected before it expired. It is satisfied
   * only if the stored token is no longer the one that failed, which means another caller already
   * replaced it and this caller should take theirs.
   */
  #tokenSatisfies(cached: EntraAccessToken | undefined, opts?: AccessTokenRequest)
      : cached is EntraAccessToken {
    if (!cached) return false;
    // Expiry gates every path — no request, however it is phrased, is answered with a token that is
    // already inside the safety window.
    if (cached.expires.valueOf() <= Date.now() + ACCESS_TOKEN_EXPIRY_SAFETY_MS) return false;
    if (opts?.staleToken !== undefined) return cached.token !== opts.staleToken;
    return !opts?.forceRefresh;
  }

  async getAccessToken(opts?: AccessTokenRequest): Promise<EntraAccessToken> {
    // Fast path, deliberately outside the lock: the overwhelmingly common case is a valid cached
    // token, and that must not serialize behind anything. It also runs before the refresh-token
    // check, because a sign-in-only grant has an access token and never has a refresh token.
    let cached = this.ctx.storage.kv.get<EntraAccessToken>("accessToken");
    if (this.#tokenSatisfies(cached, opts)) {
      return cached;
    }

    let config = this.#requireConfig();

    if (!this.ctx.storage.kv.get<string>("refreshToken")) {
      throw new Error("This Microsoft connection can no longer be refreshed.");
    }

    // Serialized so a burst of concurrent 401s collapses into one token exchange. The re-check
    // inside the lock is what does the collapsing — the lock alone would just queue the mints.
    return this.#updateCredentials(async () => {
      let fresh = this.ctx.storage.kv.get<EntraAccessToken>("accessToken");
      if (this.#tokenSatisfies(fresh, opts)) {
        return fresh;
      }

      // A mint already established that these credentials are permanently dead. Fail the same way
      // without asking Entra again — see MINT_FAILURE_COOLDOWN_MS.
      let recorded = this.ctx.storage.kv.get<StoredMintFailure>("mintFailure");
      if (recorded && Date.now() - recorded.at < MINT_FAILURE_COOLDOWN_MS) {
        throw new Error(recorded.message);
      }

      // Re-read rather than closing over the outer value: the credentials may have been replaced
      // while this call waited for the lock.
      let refreshToken = this.ctx.storage.kv.get<string>("refreshToken");
      if (!refreshToken) {
        throw new Error("This Microsoft connection can no longer be refreshed.");
      }

      // Logged before the exchange so a mint that fails or hangs still leaves a trace. The events
      // are distinct because their rates mean different things: `expiry` should tick about once per
      // token lifetime, whereas `rejected` means a token was invalidated early and the 401 retry
      // healed it — and a burst of 401s should still produce exactly one.
      logger.info("minting Microsoft access token", {
        event: opts?.staleToken !== undefined
            ? "microsoft.token.mint.rejected"
            : "microsoft.token.mint.expiry",
      });

      let result = await refreshAccessToken({
        ...config,
        refreshToken,
        scopes: this.ctx.storage.kv.get<string[]>("grantScopes") ?? IDENTITY_SCOPES,
      }, AbortSignal.timeout(TOKEN_MINT_TIMEOUT_MS));

      if (!result.ok) {
        this.#recordMintFailure(result.failure);
        throw new Error(result.failure.message);
      }

      // Backstop for the credential mutators: the mutex already keeps them from interleaving with a
      // mint, so this should be unreachable. It stays because the token we just minted was issued
      // against credentials this account may no longer hold — a reconnect that landed in between
      // would have replaced them — and publishing it would resurrect a superseded grant. Note the
      // comparison is against the token we presented, not against a token we assume is unchanged:
      // the mint itself rotates the stored value, which is why the store happens after this check.
      if (this.ctx.storage.kv.get<string>("refreshToken") !== refreshToken) {
        logger.warn("discarded a Microsoft access token minted against superseded credentials", {
          event: "microsoft.token.mint.superseded",
        });
        let current = this.ctx.storage.kv.get<EntraAccessToken>("accessToken");
        if (current) return current;
        throw new Error("Microsoft credentials changed while refreshing. Please try again.");
      }

      this.#storeGrant(result.grant);
      this.ctx.storage.kv.delete("mintFailure");
      return result.grant.accessToken;
    });
  }

  /**
   * Report that Microsoft Graph rejected this account's token in a way a new token cannot fix — a
   * 401 carrying a claims challenge.
   *
   * Under Conditional Access or continuous access evaluation the token endpoint keeps minting
   * happily while every Graph call fails, which leaves the account undead: refreshes succeed, work
   * fails, and the UI never offers a reconnect. Treating the challenge as credential death is what
   * surfaces the reconnect prompt.
   */
  async reportCredentialsRejected(detail?: string): Promise<void> {
    await this.#updateCredentials(async () => {
      // The current token is known-bad, so don't keep serving it from cache.
      this.ctx.storage.kv.delete("accessToken");
      this.#recordMintFailure({
        permanent: true,
        codes: [],
        message: "Microsoft requires this account to sign in again" +
            (detail ? ` (${detail})` : "") + ". Please reconnect the account.",
      });
    });
  }

  // Record a permanent failure and, the first time an account dies, tell the Workshop. Transient
  // failures are not recorded: the next caller should be free to try again immediately.
  #recordMintFailure(failure: TokenFailure): void {
    if (!failure.permanent) return;
    let previous = this.ctx.storage.kv.get<StoredMintFailure>("mintFailure");
    this.ctx.storage.kv.put<StoredMintFailure>("mintFailure", {
      message: failure.message, at: Date.now(),
    });
    // Recorded before notifying so only the first caller of a burst notifies.
    if (!previous) this.#notifyCredentialsDead();
  }

  // Tell the workshop the credentials are permanently dead so the UI prompts a reconnect instead of
  // every call failing opaquely. Fire and forget — a notification failure must not mask the error.
  #notifyCredentialsDead(): void {
    let callback = this.ctx.storage.kv.get<Fetcher<GatekeeperConnectCallback>>("callback");
    callback?.credentialsExpired().catch(notifyErr => {
      logger.warn("failed to notify credential expiry", {
        event: "credentials.expiry.notify.failed", error: notifyErr,
      });
    });
  }

  /**
   * The tenant's verified domains, cached per account so repeated identity checks on one account
   * cost one Graph call rather than one each.
   */
  async getVerifiedDomains(): Promise<string[]> {
    let cached = this.ctx.storage.kv.get<CachedVerifiedDomains>("verifiedDomains");
    if (cached && Date.now() - cached.fetchedAt < VERIFIED_DOMAIN_CACHE_MS) {
      return cached.domains;
    }

    let token = await this.getAccessToken();
    let domains = await fetchVerifiedDomains(token.token, AbortSignal.timeout(GRAPH_TIMEOUT_MS));
    // An empty answer is never cached: it would deny every sign-in for a day on the strength of one
    // odd response.
    if (domains.length > 0) {
      this.ctx.storage.kv.put<CachedVerifiedDomains>("verifiedDomains", {
        domains, fetchedAt: Date.now(),
      });
    }
    return domains;
  }

  async alarm(_alarmInfo?: AlarmInvocationInfo): Promise<void> {
    // Drop the account if the flow never completed, or if this was a transient sign-in-only grant
    // (used once to read the email for login). Serialized so the wipe cannot land in the middle of
    // a mint, leaving a freshly minted token behind on a deleted account.
    await this.#updateCredentials(async () => {
      if (!this.hasRefreshToken() || this.ctx.storage.kv.get<boolean>("authOnly")) {
        this.ctx.storage.deleteAll();
      }
    });
  }

  /**
   * Entra publishes no token-revocation endpoint for this flow, so disconnecting is a local delete:
   * the stored refresh token is destroyed here, and Microsoft keeps its own record of the grant
   * until it lapses. A tenant administrator can end the sessions org-side (Entra portal, or
   * `revokeSignInSessions` on the user).
   */
  async revoke(): Promise<void> {
    await this.#updateCredentials(async () => {
      this.ctx.storage.deleteAlarm();
      this.ctx.storage.deleteAll();
    });
  }
}

type GatekeeperUserImplProps = {
  userObjectId: string;
}

@validateRpc()
export class GatekeeperUserImpl extends WorkerEntrypoint<Env, GatekeeperUserImplProps>
                                implements GatekeeperUser {
  #account(): DurableObjectStub<UserAccount> {
    let id = this.ctx.exports.UserAccount.idFromString(this.ctx.props.userObjectId);
    return this.ctx.exports.UserAccount.get(id);
  }

  async describe(): Promise<AccountDescription> {
    let account = this.#account();
    let grantedResourcesPromise = account.getGrantedResourceUrlPatterns();
    let token = await account.getAccessToken();
    // Accounts without a photo are common, and a missing avatar is not a reason to fail the whole
    // description, so the logo stands in.
    let [profile, avatarUrl, grantedResourceUrlPatterns] = await Promise.all([
      fetchGraphProfile(token.token, AbortSignal.timeout(GRAPH_TIMEOUT_MS)),
      fetchAvatarDataUrl(token.token, AbortSignal.timeout(GRAPH_TIMEOUT_MS)),
      grantedResourcesPromise,
    ]);

    return {
      displayName: typeof profile.displayName === "string" ? profile.displayName : undefined,
      uniqueName: typeof profile.mail === "string" && profile.mail
          ? profile.mail
          : typeof profile.userPrincipalName === "string" ? profile.userPrincipalName : undefined,
      avatar: { url: avatarUrl ?? MICROSOFT_LOGO_URL },
      grantedResourceUrlPatterns,
    };
  }

  async getAuthenticatedEmail(): Promise<string | null> {
    // Contract is Promise<string | null>: never throw. The access token fetch can throw if the
    // (possibly transient sign-in) grant has been cleaned up, and the Graph calls can throw on a
    // non-2xx response — treat any failure as "no email available".
    try {
      let account = this.#account();
      let token = await account.getAccessToken();
      if (!token) return null;

      let [claims, verifiedDomains, profile] = await Promise.all([
        account.getIdTokenClaims(),
        account.getVerifiedDomains(),
        fetchGraphProfile(token.token, AbortSignal.timeout(GRAPH_TIMEOUT_MS)),
      ]);

      return resolveVerifiedEmail(profile, claims, verifiedDomains, this.env.TENANT_ID ?? "");
    } catch {
      return null;
    }
  }

  /**
   * Display name and photo for the Workshop to seed a new user's profile with. Hints only: the
   * sign-in identity is getAuthenticatedEmail() and nothing here feeds it.
   *
   * Same never-throw contract as getAuthenticatedEmail(). The token fetch can throw once the
   * transient sign-in grant has been cleaned up, and the profile read can throw on a non-2xx Graph
   * response; either way the provider simply has nothing to offer and sign-in continues. Absent
   * hints are omitted rather than set to undefined, so an RPC caller sees an empty object.
   */
  async getAuthenticatedProfile(): Promise<{
    name?: string;
    photo?: ProfilePhoto;
  }> {
    try {
      let token = await this.#account().getAccessToken();
      if (!token) return {};

      // The photo lookup walks several endpoints, so it runs alongside the profile read rather than
      // after it; it reports "no photo" instead of failing, which is why only one can reject here.
      let [profile, photo] = await Promise.all([
        fetchGraphProfile(token.token, AbortSignal.timeout(GRAPH_TIMEOUT_MS)),
        fetchProfilePhoto(token.token, AbortSignal.timeout(GRAPH_TIMEOUT_MS)),
      ]);

      let name = typeof profile.displayName === "string" ? profile.displayName.trim() : "";
      return { ...(name ? { name } : {}), ...(photo ? { photo } : {}) };
    } catch {
      return {};
    }
  }

  async getSupportedResources(): Promise<SupportedResource[]> {
    return SUPPORTED_RESOURCES;
  }

  async getGatekeeperClassFor(url: string): Promise<{
    class: DurableObjectClass<Gatekeeper<any>>;
    resource: SupportedResource;
  }> {
    let parsed = new URL(url);

    // Exact host match: a substring test would also accept `outlook.office.com.example.com`. The
    // path is matched on a segment boundary so a calendar, people, or `/mailbox…` URL on the same
    // host cannot bind the mailbox.
    let isMailPath = parsed.pathname === "/mail" || parsed.pathname.startsWith("/mail/");
    if (parsed.hostname === "outlook.office.com" && isMailPath) {
      let props: OutlookMailGatekeeperImplProps = { userObjectId: this.ctx.props.userObjectId };
      return {
        class: this.ctx.exports.OutlookMailGatekeeperImpl({props}),
        resource: OUTLOOK_MAIL_RESOURCE,
      };
    }

    throw new Error(`The Microsoft gatekeeper cannot connect this URL: ${url}`);
  }

  async startResourceConfigurator(
      resourceUrlPattern: string): Promise<ResourceConfiguratorFrame> {
    if (resourceUrlPattern === OUTLOOK_MAIL_RESOURCE.urlPattern) {
      return {
        iframeHtml: OUTLOOK_MAIL_CONFIGURATOR_HTML,
        ui: new RpcStub(new OutlookMailConfiguratorUI()),
      };
    }

    throw new Error(`Unsupported resource configurator type: ${resourceUrlPattern}`);
  }

  async revoke(): Promise<void> {
    await this.#account().revoke();
  }

  async reconnect(): Promise<{url: string}> {
    let account = this.#account();
    let initiationNonce = generateNonce();
    // Re-request the scopes already granted so a plain reconnect doesn't narrow access.
    await account.prepareReconnect(initiationNonce, await account.getGrantScopes());
    return { url: `${getBaseUrl(this.env)}/${this.ctx.props.userObjectId}/${initiationNonce}` };
  }

  async ensureResources(resourceUrlPatterns: string[]): Promise<{url?: string}> {
    validateResourceUrlPatterns(resourceUrlPatterns);

    let account = this.#account();
    let granted = new Set(await account.getGrantedResourceUrlPatterns());
    if (resourceUrlPatterns.every(pattern => granted.has(pattern))) {
      return {};
    }

    // Request the union of what's already granted and what's newly needed, so the expansion never
    // drops existing access.
    let unionPatterns = new Set([...granted, ...resourceUrlPatterns]);
    let requestedScopes = resourceUrlPatternsToOAuthScopes([...unionPatterns]);
    let initiationNonce = generateNonce();
    await account.prepareReconnect(initiationNonce, requestedScopes);
    return { url: `${getBaseUrl(this.env)}/${this.ctx.props.userObjectId}/${initiationNonce}` };
  }

  /**
   * Mint a verifier representing this account. The overseer mints one on every open, so it must
   * exist even while no resource type consults it.
   */
  @skipRpcValidation()
  async getVerifier(): Promise<Fetcher<GatekeeperUserVerifier>> {
    let props: MicrosoftVerifierProps = { userObjectId: this.ctx.props.userObjectId };
    return this.ctx.exports.MicrosoftVerifier({ props });
  }
}

// ---------------------------------------------------------------------------
// Verifier
//
// Opaque capability standing for "this account", handed to Gatekeeper.addObserver() and passed back
// to this same gatekeeper, which may therefore trust what it reports. It carries the account id in
// props so any future access check runs against the observer's own Microsoft token rather than the
// resource owner's. No resource type consults it yet, so it exposes no checks.

type MicrosoftVerifierProps = {
  userObjectId: string;
};

@validateRpc()
export class MicrosoftVerifier extends WorkerEntrypoint<Env, MicrosoftVerifierProps>
    implements GatekeeperUserVerifier {
}

// ---------------------------------------------------------------------------
// Resource configurator
//
// RPC interface exposed to the resource-selection iframe. The mailbox is a singleton — a connected
// account has exactly one — so the frame confirms the choice and needs nothing from the gatekeeper.

@validateRpc()
export class OutlookMailConfiguratorUI extends RpcTarget implements OutlookMailConfiguratorRpc {}
