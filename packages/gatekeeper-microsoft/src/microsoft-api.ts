// Helpers talking to Microsoft Entra ID (OAuth 2.0 v2 endpoints) and Microsoft Graph.
//
// Everything here is plain `fetch`: the identity library (MSAL) is built for interactive clients
// and brings a large dependency tree for two endpoints we call by hand.
//
// Nothing in this file may log tokens. Error text from Entra is safe to surface (it carries AADSTS
// codes, not credentials) but is length-capped before it reaches a message a user can see.

const LOGIN_HOST = "https://login.microsoftonline.com";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/** Longest error text we echo from a provider response into an error message. */
const MAX_PROVIDER_ERROR_CHARS = 500;

/** Ceiling on a profile photo we are willing to inline as a data URI. */
const MAX_AVATAR_BYTES = 256 * 1024;

/**
 * Ceiling on a photo handed to the Workshop as a sign-in profile hint, matching the size the
 * Workshop accepts for a stored avatar. A Worker cannot re-encode an image, so a photo over the
 * limit is answered by asking Graph for a smaller rendition, never by transcoding this one.
 */
const MAX_PROFILE_PHOTO_BYTES = 100 * 1024;

/** Longest address we accept as a sign-in identity (RFC 5321 practical maximum). */
const MAX_EMAIL_CHARS = 320;

export type EntraAccessToken = {
  /** Never log this. */
  token: string;
  expires: Date;
};

/**
 * Options for requesting an access token.
 *
 * `forceRefresh` means "do not serve me a cached token" — the caller saw a 401, so a token cached
 * against its recorded expiry is known-bad. `staleToken` names the rejected token so a burst of
 * concurrent 401s collapses into a single token exchange instead of one exchange per caller.
 */
export type AccessTokenRequest = {
  forceRefresh?: boolean;
  /** The token the caller just had rejected. Never log this. */
  staleToken?: string;
};

/**
 * The id_token claims we retain.
 *
 * `tid` is the tenant the signing-in account belongs to; `oid` is the account's immutable object id
 * within that tenant. Both are read from the id_token returned by the token endpoint — see
 * `decodeIdTokenClaims` for why that source, and only that source, may be trusted unverified.
 */
export type IdTokenClaims = {
  tid?: string;
  oid?: string;
};

export type EntraTokenGrant = {
  accessToken: EntraAccessToken;
  /**
   * Entra rotates refresh tokens: every mint returns a new one and retires the one presented, so a
   * caller that keeps the old value loses the account when the retired token stops working.
   * Undefined when the grant did not include `offline_access` (sign-in-only grants).
   */
  refreshToken?: string;
  /** Scopes Entra reports as granted, exactly as returned (resource-qualified for Graph scopes). */
  grantedScopes: string[];
  idTokenClaims?: IdTokenClaims;
};

/**
 * A token-endpoint rejection, classified by whether asking again could ever succeed.
 *
 * `permanent` failures mean the connection is dead until a human acts (reconnect, admin consent, a
 * rotated client secret); they suppress further mints and tell the Workshop to offer a reconnect.
 * Everything else is transient — an unknown code is treated as transient so a code we have not
 * catalogued cannot silently kill a live account.
 */
export type TokenFailure = {
  permanent: boolean;
  codes: number[];
  /** User-facing explanation. Safe to display: contains no credentials. */
  message: string;
};

export type TokenMintResult =
  | { ok: true; grant: EntraTokenGrant }
  | { ok: false; failure: TokenFailure };

/**
 * AADSTS codes that no retry can fix.
 *
 * Everything here needs either a fresh user consent/sign-in or an administrator action, so
 * continuing to hit the token endpoint only burns rate limit while the account looks alive.
 */
const PERMANENT_AADSTS_CODES = new Set([
  70008,   // authorization code or refresh token expired
  50173,   // grant invalidated (password change, admin session revocation)
  65001,   // the user or administrator has not consented to the requested permissions
  700082,  // refresh token expired due to inactivity
  7000215, // invalid client secret
  7000222, // client secret expired
]);

/** Codes meaning the deployment's own client secret is wrong or lapsed, not the user's grant. */
const CLIENT_SECRET_AADSTS_CODES = new Set([7000215, 7000222]);

/**
 * OAuth error identifiers that, by definition, cannot be resolved without the user going through an
 * interactive sign-in — Conditional Access step-up, an MFA challenge, an expired session, or a
 * consent prompt.
 *
 * Classifying on the identifier rather than the numeric code matters: the code space behind these
 * (50076, 50079, 50097, 530003, …) is large, policy-dependent, and grows, so a refresh loop would
 * otherwise retry forever on an account only a reconnect can revive, and the Workshop would never
 * be told to offer one.
 */
const INTERACTION_REQUIRED_ERRORS = new Set([
  "interaction_required",
  "login_required",
  "consent_required",
]);

function truncate(value: string, max: number = MAX_PROVIDER_ERROR_CHARS): string {
  return value.length > max ? `${value.slice(0, max)}…` : value;
}

export function authorizeEndpoint(tenantId: string): string {
  return `${LOGIN_HOST}/${encodeURIComponent(tenantId)}/oauth2/v2.0/authorize`;
}

export function tokenEndpoint(tenantId: string): string {
  return `${LOGIN_HOST}/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
}

/**
 * Classify a token-endpoint error body.
 *
 * Entra reports the specific cause twice: as `error_codes` (numeric) and embedded in
 * `error_description` as `AADSTS<code>`. Both are read, because the array is absent from some error
 * shapes (notably certain HTTP-level failures) while the description is present. The top-level
 * OAuth `error` identifier is read as well, since it classifies the whole interaction-required
 * family without enumerating its codes.
 */
export function classifyTokenFailure(body: unknown): TokenFailure {
  const record = (body ?? {}) as {
    error?: unknown; error_description?: unknown; error_codes?: unknown;
  };
  const error = typeof record.error === "string" ? record.error : "";
  const description = typeof record.error_description === "string" ? record.error_description : "";

  const codes = new Set<number>();
  if (Array.isArray(record.error_codes)) {
    for (const code of record.error_codes) {
      if (typeof code === "number" && Number.isFinite(code)) codes.add(code);
    }
  }
  for (const match of description.matchAll(/AADSTS(\d{4,10})/g)) {
    codes.add(Number(match[1]));
  }

  const codeList = [...codes];
  const interactionRequired = INTERACTION_REQUIRED_ERRORS.has(error);
  const secretProblem = codeList.some(code => CLIENT_SECRET_AADSTS_CODES.has(code));
  const permanent = interactionRequired || codeList.some(code => PERMANENT_AADSTS_CODES.has(code));
  const detail = truncate(description || error || "no detail provided");

  let message: string;
  if (secretProblem) {
    message = "This Cloudflare OS instance's Microsoft client secret is invalid or has expired. " +
        `An administrator must rotate it — reconnecting will not help (${detail})`;
  } else if (interactionRequired) {
    message = `Microsoft requires you to sign in again to satisfy an access policy (${detail}). ` +
        "Please reconnect the account.";
  } else if (permanent) {
    message = `Microsoft credentials have expired or been revoked (${detail}). ` +
        "Please reconnect the account.";
  } else {
    message = `Microsoft rejected the token request (${detail}). Please try again.`;
  }

  return { permanent, codes: codeList, message };
}

/**
 * Prefix Entra puts in front of a Graph permission when it reports granted scopes in resource-
 * qualified form.
 */
const GRAPH_RESOURCE_PREFIX = "https://graph.microsoft.com/";

/**
 * Reserved OpenID Connect scopes. Entra never lists these in a token response's `scope`, so a
 * grant that asked for them must be treated as having them.
 */
const IMPLIED_OIDC_SCOPES = new Set(["openid", "profile", "email", "offline_access"]);

/**
 * Reduce a scope string to the form scopes are compared in.
 *
 * Entra's token response does not echo the scopes we requested: a Graph permission may come back
 * resource-qualified (`https://graph.microsoft.com/Mail.ReadWrite`) or in different casing
 * (`Mail.Readwrite`), and reserved OIDC scopes are omitted entirely. Comparing the raw strings
 * would decide a permission was never granted, so every caller would ask for consent again — and
 * the next response would look exactly the same, forever.
 */
export function normalizeScope(scope: string): string {
  let normalized = scope.trim().toLowerCase();
  return normalized.startsWith(GRAPH_RESOURCE_PREFIX)
      ? normalized.slice(GRAPH_RESOURCE_PREFIX.length)
      : normalized;
}

/** Whether a grant reporting `grantedScopes` covers everything in `requiredScopes`. */
export function grantCoversScopes(grantedScopes: string[], requiredScopes: string[]): boolean {
  const granted = new Set(
      grantedScopes.map(normalizeScope).filter(Boolean));
  return requiredScopes.every(scope => {
    const normalized = normalizeScope(scope);
    return !normalized || IMPLIED_OIDC_SCOPES.has(normalized) || granted.has(normalized);
  });
}

/**
 * Read the claims out of an id_token WITHOUT verifying its signature.
 *
 * This is sound only because every id_token reaching this function came straight from the
 * tenant-pinned token endpoint over TLS in the response to our own client-authenticated request:
 * the transport is the proof of origin. An id_token arriving any other way — a redirect query
 * string, a fragment, a request body, an API response — is attacker-supplied and MUST NOT be fed
 * here. The authorization request accordingly asks for `response_type=code` only, so no id_token
 * ever crosses the front channel.
 */
export function decodeIdTokenClaims(idToken: string | undefined): IdTokenClaims | undefined {
  if (!idToken) return undefined;
  const parts = idToken.split(".");
  if (parts.length !== 3) return undefined;
  try {
    const json = atob(parts[1].replaceAll("-", "+").replaceAll("_", "/"));
    const claims = JSON.parse(json) as { tid?: unknown; oid?: unknown };
    return {
      tid: typeof claims.tid === "string" ? claims.tid : undefined,
      oid: typeof claims.oid === "string" ? claims.oid : undefined,
    };
  } catch {
    return undefined;
  }
}

function grantFromTokenResponse(body: {
  access_token?: unknown; expires_in?: unknown; refresh_token?: unknown;
  scope?: unknown; id_token?: unknown;
}): EntraTokenGrant {
  if (typeof body.access_token !== "string" || !body.access_token) {
    throw new Error("The Microsoft token endpoint returned no access token.");
  }
  // Entra always reports `expires_in`; fall back to a conservative hour rather than treating a
  // missing value as an immediately-expired token.
  const expiresInSeconds = typeof body.expires_in === "number" && body.expires_in > 0
      ? body.expires_in
      : 3600;
  return {
    accessToken: {
      token: body.access_token,
      expires: new Date(Date.now() + expiresInSeconds * 1000),
    },
    refreshToken: typeof body.refresh_token === "string" && body.refresh_token
        ? body.refresh_token
        : undefined,
    grantedScopes: typeof body.scope === "string" ? body.scope.split(" ").filter(Boolean) : [],
    idTokenClaims: decodeIdTokenClaims(
        typeof body.id_token === "string" ? body.id_token : undefined),
  };
}

async function readTokenErrorBody(response: Response): Promise<unknown> {
  const contentType = response.headers.get("Content-Type") ?? "";
  if (contentType.includes("application/json")) {
    return await response.json().catch(() => undefined);
  }
  const text = await response.text().catch(() => "");
  return { error_description: text || `${response.status} ${response.statusText}` };
}

export type TokenRequestConfig = {
  tenantId: string;
  clientId: string;
  clientSecret: string;
};

/**
 * Redeem an authorization code. `signal` bounds the round trip; the caller holds the credential
 * mutex across it.
 */
export async function exchangeAuthCode(
    config: TokenRequestConfig & { code: string; redirectUri: string; scopes: string[] },
    signal?: AbortSignal): Promise<EntraTokenGrant> {
  const params = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    code: config.code,
    redirect_uri: config.redirectUri,
    grant_type: "authorization_code",
    scope: config.scopes.join(" "),
  });

  const response = await fetch(tokenEndpoint(config.tenantId), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params,
    ...(signal ? { signal } : {}),
  });

  if (!response.ok) {
    throw new Error(classifyTokenFailure(await readTokenErrorBody(response)).message);
  }
  return grantFromTokenResponse(await response.json());
}

/**
 * Exchange a refresh token for a new access token.
 *
 * The returned grant carries a NEW refresh token, which the caller must persist in place of the one
 * it presented. `scopes` re-states the scopes the account consented to; Entra requires the refresh
 * request to name a subset of the original grant.
 */
export async function refreshAccessToken(
    config: TokenRequestConfig & { refreshToken: string; scopes: string[] },
    signal?: AbortSignal): Promise<TokenMintResult> {
  const params = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    refresh_token: config.refreshToken,
    grant_type: "refresh_token",
    scope: config.scopes.join(" "),
  });

  const response = await fetch(tokenEndpoint(config.tenantId), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params,
    ...(signal ? { signal } : {}),
  });

  if (!response.ok) {
    return { ok: false, failure: classifyTokenFailure(await readTokenErrorBody(response)) };
  }
  return { ok: true, grant: grantFromTokenResponse(await response.json()) };
}

// =======================================================================================
// Microsoft Graph
// =======================================================================================

/** The `/me` fields the identity checks depend on. Values are untrusted until validated. */
export type GraphUserProfile = {
  id?: unknown;
  displayName?: unknown;
  mail?: unknown;
  userPrincipalName?: unknown;
  userType?: unknown;
};

async function graphGet(path: string, accessToken: string, signal?: AbortSignal): Promise<Response> {
  return await fetch(`${GRAPH_BASE}${path}`, {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" },
    ...(signal ? { signal } : {}),
  });
}

async function graphGetJson(
    path: string, accessToken: string, signal?: AbortSignal): Promise<unknown> {
  const response = await graphGet(path, accessToken, signal);
  if (!response.ok) {
    await response.body?.cancel();
    throw new Error(`Microsoft Graph request failed: ${response.status} ${response.statusText}`);
  }
  return await response.json();
}

/**
 * Read the signed-in user's profile.
 *
 * `userType` is selected explicitly and is load-bearing: without it a B2B guest is
 * indistinguishable from a tenant member (see `resolveVerifiedEmail`).
 */
export async function fetchGraphProfile(
    accessToken: string, signal?: AbortSignal): Promise<GraphUserProfile> {
  const body = await graphGetJson(
      "/me?$select=id,displayName,mail,userPrincipalName,userType", accessToken, signal);
  return (body ?? {}) as GraphUserProfile;
}

/**
 * Read the tenant's verified domains.
 *
 * Graph exposes no per-address verification flag, so "the address is on a domain this tenant has
 * proven it owns" is what makes an address provider-verified. `User.Read` is enough to read
 * `verifiedDomains` on the organization; every other organization property comes back null under
 * that permission, which is fine — this is the only one wanted.
 */
export async function fetchVerifiedDomains(
    accessToken: string, signal?: AbortSignal): Promise<string[]> {
  const body = await graphGetJson("/organization?$select=verifiedDomains", accessToken, signal) as {
    value?: unknown;
  };
  const organizations = Array.isArray(body.value) ? body.value : [];
  const domains = new Set<string>();
  for (const organization of organizations) {
    const entries = (organization as { verifiedDomains?: unknown }).verifiedDomains;
    if (!Array.isArray(entries)) continue;
    for (const entry of entries) {
      const name = (entry as { name?: unknown }).name;
      if (typeof name === "string" && name.trim()) domains.add(name.trim().toLowerCase());
    }
  }
  return [...domains];
}

/**
 * Image types allowed to appear in an avatar data URI.
 *
 * The media type is copied into a URL that the Workshop renders, so it is an allowlist rather than
 * a sanitizer: anything Graph reports that is not a plain raster image (a rewritten proxy, a
 * document type, `image/svg+xml` with its scripting surface) is treated exactly like "no photo".
 */
const ALLOWED_AVATAR_TYPES = new Set([
  "image/jpeg", "image/png", "image/gif", "image/webp",
]);

type AvatarMediaType = "image/jpeg" | "image/png" | "image/gif" | "image/webp";

function isAllowedAvatarType(mediaType: string): mediaType is AvatarMediaType {
  return ALLOWED_AVATAR_TYPES.has(mediaType);
}

/**
 * Media types a sign-in profile photo may carry — narrower than the avatar allowlist because the
 * Workshop stores these bytes as the user's avatar and accepts only JPEG and PNG. A GIF or WebP
 * photo is therefore dropped here rather than failing validation inside the sign-in callback.
 */
type ProfilePhotoType = "image/jpeg" | "image/png";

function isProfilePhotoType(mediaType: string): mediaType is ProfilePhotoType {
  return mediaType === "image/jpeg" || mediaType === "image/png";
}

/** Reduce a Content-Type header to its bare media type, or null when there is none. */
function mediaTypeOf(header: string | null): string | null {
  if (!header) return null;
  // Strip any parameters (`image/jpeg; charset=…`) before matching.
  return header.split(";")[0].trim().toLowerCase() || null;
}

/**
 * Fetch an image endpoint, or null for anything that is not a usable image: a non-2xx response
 * (404 — "this account has no photo" — is the common case), a media type the caller does not
 * accept, an empty body, or a body over `maxBytes`. Never throws.
 */
async function fetchGraphImage<T extends string>(
    path: string,
    accessToken: string,
    accepts: (mediaType: string) => mediaType is T,
    maxBytes: number,
    signal?: AbortSignal): Promise<{ data: Uint8Array; mediaType: T } | null> {
  try {
    const response = await graphGet(path, accessToken, signal);
    if (!response.ok) {
      await response.body?.cancel();
      return null;
    }
    const mediaType = mediaTypeOf(response.headers.get("Content-Type"));
    if (!mediaType || !accepts(mediaType)) {
      await response.body?.cancel();
      return null;
    }
    // The original photo is served at whatever size it was uploaded — megabytes, potentially — and
    // this runs inside a sign-in. A declared length over the limit is refused before the body is
    // read, so an unusable photo costs headers instead of a download. A response that declares no
    // length is still read and measured below.
    const declaredLength = response.headers.get("Content-Length");
    if (declaredLength !== null && Number(declaredLength) > maxBytes) {
      await response.body?.cancel();
      return null;
    }
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength === 0 || buffer.byteLength > maxBytes) return null;
    return { data: new Uint8Array(buffer), mediaType };
  } catch {
    return null;
  }
}

function toBase64(bytes: Uint8Array): string {
  // Chunked so a large photo cannot blow the argument limit of String.fromCharCode.
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

/**
 * Fetch the user's profile photo as a data URI, or null if there isn't one.
 *
 * Accounts without a photo are the common case (Graph answers 404), so this never throws: the
 * caller falls back to the vendor logo rather than failing the whole account description. An
 * unrecognized media type takes that same fallback.
 */
export async function fetchAvatarDataUrl(
    accessToken: string, signal?: AbortSignal): Promise<string | null> {
  const image = await fetchGraphImage(
      "/me/photos/48x48/$value", accessToken, isAllowedAvatarType, MAX_AVATAR_BYTES, signal);
  return image && `data:${image.mediaType};base64,${toBase64(image.data)}`;
}

export type ProfilePhoto = {
  data: Uint8Array;
  mimeType: ProfilePhotoType;
};

/**
 * Photo endpoints tried, in order, for a sign-in profile hint.
 *
 * The 240px rendition comes first because the Workshop renders an avatar at 256px, so a 48px
 * connector-chip thumbnail would be visibly soft. The original is tried next: sized renditions are
 * not served for every mailbox topology, and an account with a photo may answer 404 for every size
 * while `/me/photo/$value` works. The small rendition is the last rung, reached when the original is
 * bigger than the Workshop accepts — a rendition Graph produces is also a chance at JPEG or PNG when
 * the uploaded original is in a format we cannot pass on.
 */
const PROFILE_PHOTO_PATHS = [
  "/me/photos/240x240/$value",
  "/me/photo/$value",
  "/me/photos/96x96/$value",
];

/**
 * Fetch a photo of the signed-in user for the Workshop to seed their avatar with, or null when this
 * account has none we can use.
 *
 * Like `fetchAvatarDataUrl`, this never throws: it feeds a sign-in that must complete whether or not
 * the provider has a photo to offer. Each rung is tried in turn and anything unusable — no photo at
 * that size, a type the Workshop rejects, or a body over the size limit — falls through to the next.
 */
export async function fetchProfilePhoto(
    accessToken: string, signal?: AbortSignal): Promise<ProfilePhoto | null> {
  for (const path of PROFILE_PHOTO_PATHS) {
    const image = await fetchGraphImage(
        path, accessToken, isProfilePhotoType, MAX_PROFILE_PHOTO_BYTES, signal);
    if (image) return { data: image.data, mimeType: image.mediaType };
  }
  return null;
}

// =======================================================================================
// Sign-in identity
// =======================================================================================

/**
 * Decide which address, if any, may stand as this account's sign-in identity.
 *
 * The Workshop keys user accounts by this string and re-validates nothing downstream, so every
 * control that keeps one person from signing in as another lives here. In order of how much weight
 * each carries:
 *
 *  1. Members only. A B2B guest is a real principal in this tenant whose `mail` is usually the
 *     external address they were invited by — accepting it would let anyone with an invitation
 *     claim an identity at another provider's domain.
 *  2. `#EXT#` rejection. Guest UPNs are minted as `user_external.com#EXT#@tenant.onmicrosoft.com`;
 *     rejecting that marker outright covers directory states where `userType` disagrees with how
 *     the principal was created.
 *  3. Verified-domain constraint. `mail` is an admin-writable attribute and is not guaranteed to
 *     sit on this tenant, so an address only counts as provider-verified when its domain is one the
 *     tenant has proven it owns.
 *  4. Tenant pinning. Defense in depth only: every attack above happens inside the correct tenant,
 *     so `tid` cannot be the control that stops them — it stops the different mistake of pointing
 *     the deployment at one tenant while accepting tokens minted for another.
 *
 * The result is lowercased because UPNs are case-insensitive at Microsoft while the Workshop's
 * account keys are exact strings — two casings of one address must not become two accounts.
 *
 * Returns null on any failure; callers treat null as "this account cannot sign anyone in".
 */
export function resolveVerifiedEmail(
    profile: GraphUserProfile | undefined,
    idTokenClaims: IdTokenClaims | undefined,
    verifiedDomains: string[],
    expectedTid: string): string | null {
  if (!profile) return null;

  const expected = expectedTid.trim().toLowerCase();
  const tid = typeof idTokenClaims?.tid === "string" ? idTokenClaims.tid.trim().toLowerCase() : "";
  if (!expected || tid !== expected) return null;

  const userType = typeof profile.userType === "string" ? profile.userType.trim() : "";
  if (userType.toLowerCase() !== "member") return null;

  const upn = typeof profile.userPrincipalName === "string" ? profile.userPrincipalName : "";
  const mail = typeof profile.mail === "string" ? profile.mail : "";
  if (upn.toLowerCase().includes("#ext#")) return null;

  const candidate = (mail || upn).trim().toLowerCase();
  if (!candidate || candidate.length > MAX_EMAIL_CHARS) return null;
  if (candidate.includes("#ext#") || /\s/.test(candidate)) return null;

  const at = candidate.indexOf("@");
  if (at <= 0 || at !== candidate.lastIndexOf("@") || at === candidate.length - 1) return null;

  const verified = new Set(verifiedDomains
      .filter((domain): domain is string => typeof domain === "string")
      .map(domain => domain.trim().toLowerCase())
      .filter(Boolean));
  if (!verified.has(candidate.slice(at + 1))) return null;

  return candidate;
}
