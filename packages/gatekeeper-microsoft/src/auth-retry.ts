// Shared fetch helper for the Microsoft Graph API client.
//
// Three retry/failure concerns compose here:
//
//   1. Auth (401). The access token we send can be stale or invalidated before its recorded expiry
//      — e.g. after a re-consent, an explicit revoke-then-reconnect, or the refresh-token rotation
//      Entra performs on every mint. The token is cached in the gatekeeper Durable Object's memory
//      and in the account's storage, so a stale-but-unexpired token would otherwise be served on
//      every request and the 401 could not self-heal. On a 401 we mint a fresh token via
//      `getAccessToken({ forceRefresh: true })` and retry exactly once. This is safe on any HTTP
//      method: a 401 is rejected before the request takes effect, so a write is never applied twice.
//      A 403 is an insufficient-permission error that a fresh token cannot fix, so it is never
//      retried.
//
//   2. Claims challenges (401 + `WWW-Authenticate: ... claims=...`). Conditional Access and
//      continuous access evaluation reject a token whose claims no longer satisfy tenant policy.
//      Minting again cannot help: the refresh token produces another token with the same claims, so
//      a retry would burn a token exchange and fail identically. Worse, the account looks alive —
//      refreshes succeed while every Graph call fails — so the UI never offers a reconnect. These
//      are therefore reported to the account authority as credential death (which is what surfaces
//      the reconnect prompt) and returned to the caller unretried.
//
//   3. Transient failures (429 / 5xx / network timeout). Retried with exponential backoff plus full
//      jitter, honoring `Retry-After` when present. A 5xx and a network timeout leave it ambiguous
//      whether the server already applied a write, so those replay idempotent GETs only. A 429 is
//      different: Graph rejects a throttled request before processing it, so replaying it — for any
//      method with a replayable body — cannot duplicate a write.
//
// All concerns share a single attempt counter, so the worst case is a predictable `retries + 1`
// requests: `retries` transient attempts plus at most one extra for the one-shot 401 refresh.

/**
 * Options for requesting an access token.
 *
 * `forceRefresh` means "do not serve me a cached token" — the caller saw a 401, so any token cached
 * against its recorded expiry is known-bad. `staleToken` names the rejected token so a burst of
 * concurrent 401s collapses into a single token exchange instead of one exchange per caller.
 */
export type AccessTokenRequest = {
  forceRefresh?: boolean;
  /** The token the caller just had rejected. Never log this. */
  staleToken?: string;
};

export type AccessTokenProvider = (opts?: AccessTokenRequest) => Promise<string>;

/**
 * Told that Graph rejected the current credentials in a way no fresh token can fix. Implementations
 * mark the account dead so the Workshop offers a reconnect; failures are swallowed by the caller so
 * a notification problem cannot mask the underlying request error.
 */
export type CredentialsRejectedReporter = (detail: string) => Promise<void>;

export type FetchWithAuthRetryOptions = {
  /**
   * Total attempts for transient failures, including the first. Defaults to 3. The one-shot 401
   * refresh does not consume this budget, so the worst case is `retries + 1` requests.
   */
  retries?: number;
  /** Per-attempt abort timeout in milliseconds. Omitted means no timeout is imposed. */
  timeoutMs?: number;
  /** Invoked once when a 401 carries a claims challenge. */
  onCredentialsRejected?: CredentialsRejectedReporter;
};

const BASE_DELAY_MS = 500;
const MAX_DELAY_MS = 10_000;

/** Longest value copied out of a `WWW-Authenticate` challenge into a message. */
const MAX_CHALLENGE_DETAIL_CHARS = 100;

/**
 * The short reason from a `WWW-Authenticate` claims challenge, or null when the header carries
 * none.
 *
 * Graph answers a policy rejection with `Bearer ... error="insufficient_claims", claims="<base64>"`.
 * Either marker identifies it. The `claims` blob itself is never surfaced: it is an opaque policy
 * demand meant for a token request, not for a human, and it can be kilobytes long.
 */
export function claimsChallengeDetail(header: string | null | undefined): string | null {
  if (!header) return null;
  let hasClaims = /(^|[\s,])claims\s*=/i.test(header);
  let insufficientClaims = /(^|[\s,])error\s*=\s*"?insufficient_claims/i.test(header);
  if (!hasClaims && !insufficientClaims) return null;
  let error = header.match(/(^|[\s,])error\s*=\s*"([^"]*)"/i)?.[2];
  return (error || "insufficient_claims").slice(0, MAX_CHALLENGE_DETAIL_CHARS);
}

/**
 * Whether a transient failure status is worth replaying for this method.
 *
 * A 5xx does not tell us whether the request took effect before the response, so it is only
 * replayed for idempotent GETs. A 429 means the request was throttled — rejected before Graph
 * processed it — so replaying it is safe for any method.
 */
function canRetry(status: number, method: string): boolean {
  if (status === 429) return true;
  if (status >= 500 && status <= 599) return method === "GET";
  return false;
}

function backoffDelayMs(attempt: number, retryAfter: string | null): number {
  // Prefer the server's Retry-After when present, capped. Only the delta-seconds form is parsed;
  // the HTTP-date form yields NaN and falls through to exponential backoff, which is fine.
  if (retryAfter) {
    let seconds = parseInt(retryAfter, 10);
    if (!Number.isNaN(seconds)) return Math.min(seconds * 1000, MAX_DELAY_MS);
  }
  // Exponential backoff with full jitter.
  let capped = Math.min(BASE_DELAY_MS * 2 ** attempt, MAX_DELAY_MS);
  return Math.random() * capped;
}

/**
 * Perform an authenticated fetch, injecting a Bearer token and applying the concerns described at
 * the top of this file. The `Authorization` header is set from the (possibly refreshed) token on
 * every attempt, so callers must NOT set it themselves; any other headers in `init.headers` are
 * preserved.
 *
 * The response body is never consumed on the path that returns, so the caller can always read the
 * error payload.
 */
export async function fetchWithAuthRetry(
  url: string,
  init: RequestInit,
  getAccessToken: AccessTokenProvider,
  opts: FetchWithAuthRetryOptions = {},
): Promise<Response> {
  let method = (init.method ?? "GET").toUpperCase();
  let retries = opts.retries ?? 3;

  // A request can only be replayed if its body can be sent again. A string body (what every call
  // site uses today) re-serializes fine; a stream is consumed by the first attempt, so retrying it
  // would send an empty or errored body. Nothing to replay is likewise fine.
  let replayable = init.body === undefined || init.body === null || typeof init.body === "string";

  // One-shot: a 401 buys exactly one refreshed retry
  let refreshed = false;
  let token = await getAccessToken();
  let attempt = 0;

  while (true) {
    // A fresh timeout signal per attempt so a retry gets the full budget, combined with any
    // caller-supplied signal.
    let signals: AbortSignal[] = [];
    if (opts.timeoutMs !== undefined) signals.push(AbortSignal.timeout(opts.timeoutMs));
    if (init.signal) signals.push(init.signal);
    let signal = signals.length > 0 ? AbortSignal.any(signals) : undefined;

    let headers = new Headers(init.headers);
    headers.set("Authorization", `Bearer ${token}`);

    let response: Response;
    try {
      response = await fetch(url, {
        ...init,
        headers,
        ...(signal ? { signal } : {}),
      });
    } catch (error) {
      // Network error or timeout: ambiguous, so retry idempotent GETs only.
      if (replayable && method === "GET" && attempt < retries - 1) {
        await new Promise(resolve => setTimeout(resolve, backoffDelayMs(attempt, null)));
        attempt++;
        continue;
      }
      throw error;
    }

    if (response.status === 401) {
      let claimsDetail = claimsChallengeDetail(response.headers.get("WWW-Authenticate"));
      if (claimsDetail) {
        // Awaited so the account is recorded dead before the caller's error propagates: the two
        // must not race, or the user sees "request failed" while the UI still believes the
        // connection is healthy. A reporting failure is swallowed — it must not replace the real
        // error with a notification error.
        if (opts.onCredentialsRejected) {
          await opts.onCredentialsRejected(claimsDetail).catch(() => {});
        }
        return response;
      }

      if (!refreshed && replayable) {
        // Deliberately does not touch `attempt`: the refresh is one-shot, so it can add at most one
        // request to the budget rather than doubling it.
        refreshed = true;
        await response.body?.cancel();
        // Naming the rejected token lets the authority collapse a concurrent burst of 401s into a
        // single token exchange — see AccessTokenRequest.
        token = await getAccessToken({ forceRefresh: true, staleToken: token });
        continue;
      }
    }

    if (replayable && canRetry(response.status, method) && attempt < retries - 1) {
      let delay = backoffDelayMs(attempt, response.headers.get("Retry-After"));
      await response.body?.cancel();
      await new Promise(resolve => setTimeout(resolve, delay));
      attempt++;
      continue;
    }

    return response;
  }
}

/**
 * How far ahead of an access token's recorded expiry we treat it as already expired, so a token
 * doesn't lapse mid-request. Must agree with the `UserAccount`'s own safety window — otherwise the
 * caching layer happily serves a token the authoritative layer would already have replaced.
 */
export const ACCESS_TOKEN_EXPIRY_SAFETY_MS = 60 * 1000;

/** Mints an access token, bypassing any cache of its own when `forceRefresh` is set. */
export type MintedAccessToken = { token: string; expires: Date };

export type MintAccessToken = (opts?: AccessTokenRequest) => Promise<MintedAccessToken>;

/**
 * Per-Durable-Object memo of the access token minted by the `UserAccount`, so a gatekeeper doesn't
 * pay an RPC on every Graph call. The token is re-minted once it is within `skewMs` of expiring, and
 * `forceRefresh` discards it outright — which is what makes the 401 retry in `fetchWithAuthRetry`
 * able to heal a token Entra invalidated ahead of its recorded expiry.
 *
 * Each gatekeeper Durable Object holds its own instance, so these can briefly diverge: two
 * gatekeepers may sit on different vintages of a token, both valid. They converge because every miss
 * goes to the same `UserAccount`, which is the single authority and the only thing that mints.
 */
export class AccessTokenCache {
  #cached: MintedAccessToken | undefined;
  #mint: MintAccessToken;
  #skewMs: number;

  constructor(mint: MintAccessToken, skewMs: number = ACCESS_TOKEN_EXPIRY_SAFETY_MS) {
    this.#mint = mint;
    this.#skewMs = skewMs;
  }

  /**
   * Whether the memoized token can answer this request.
   *
   * The `staleToken` arm mirrors the re-check the authority performs, one layer up: a caller whose
   * token was just rejected can be served locally if this cache has already moved past that token,
   * because some earlier caller in the same 401 burst already replaced it.
   */
  #satisfies(cached: MintedAccessToken | undefined, opts?: AccessTokenRequest)
      : cached is MintedAccessToken {
    if (!cached) return false;
    if (cached.expires.valueOf() <= Date.now() + this.#skewMs) return false;
    if (opts?.staleToken !== undefined) return cached.token !== opts.staleToken;
    return !opts?.forceRefresh;
  }

  async get(opts?: AccessTokenRequest): Promise<string> {
    let cached = this.#cached;
    if (!this.#satisfies(cached, opts)) {
      cached = await this.#mint(opts);
      this.#cached = cached;
    }
    return cached.token;
  }

  /**
   * Forget the memoized token, so the next `get()` asks the authority again.
   *
   * Needed for the claims-challenge path, which deliberately does not pass `forceRefresh`: a
   * re-mint before the user reconnects would produce another token carrying the same rejected
   * claims. Without this, the rejected token would keep being served from memory until its recorded
   * expiry — including after a successful reconnect, so every call would report the account dead
   * again and the reconnect prompt would keep reappearing. Dropping it locally means the first call
   * after a reconnect reaches the authority and gets the new grant (or the recorded failure, if the
   * account really is still dead).
   */
  invalidate(): void {
    this.#cached = undefined;
  }
}
