// Sign-in via authentication gatekeepers.
//
// Unlike the normal connect-account flow (which runs for an already-logged-in user), login happens
// before we know who the user is. The PublicApi starts a gatekeeper connect flow (in "auth" scope
// mode) with a `LoginConnectCallbackImpl` as the callback and a `PendingLogin` DO to bridge the
// result back to the waiting browser:
//
//   1. PublicApi.startGatekeeperLogin(vendorId) creates a PendingLogin DO (keyed by a random DO id),
//      hands the gatekeeper a LoginConnectCallbackImpl, and returns {url, attempt}, where `attempt`
//      is an RpcStub wrapping the DO (so the client awaits via a capability, never a guessable id).
//   2. The browser opens `url` (the gatekeeper's self-closing OAuth popup) and calls
//      `attempt.wait()`, which blocks on the PendingLogin DO.
//   3. When the gatekeeper finishes, it calls LoginConnectCallbackImpl.complete(user). We read the
//      verified email, resolve/create the email-keyed user DO, mint a session, and deliver the token
//      to the PendingLogin DO, which resolves the awaiting RPC.
//
// Sign-in only requests minimal scopes and the gatekeeper grant is transient (it self-destructs
// shortly after we read the email) — so login does NOT create a persistent connected account.
// Capability access (repos, docs, billing) is granted later when the user explicitly connects the
// gatekeeper, which requests the full scopes and persists the connection.

import { DurableObject, WorkerEntrypoint } from "cloudflare:workers";
import { GatekeeperConnectCallback, GatekeeperUser } from "@gadgets/workshop-shared/gatekeeper";
import { createWorkshopLogger } from "../observability";
import { CLOUDFLARE_VENDOR_ID, sanitizeDisplayNameHint } from "../user.js";
import { readAdminConfig } from "../admin-config.js";
import { AvatarEnv, putUserAvatar, userAvatarExists, validateAvatarBytes } from "../avatar.js";

const logger = createWorkshopLogger("workshop.auth");

type LoginLogger = ReturnType<typeof createWorkshopLogger>;

// How long sign-in waits for the gatekeeper's profile hints. Deliberately much tighter than the
// timeouts a gatekeeper applies to its own provider calls: the hints are a nicety, so a slow
// provider costs the user a seeded name, never a slow sign-in.
const PROFILE_HINT_TIMEOUT_MS = 5000;

// getAuthenticatedProfile() is optional on GatekeeperUser. We don't probe the stub for it — the
// vendor's providesAuthProfile declaration is what we gate on — but TypeScript still can't call an
// optional method on the stub type, so we view the stub through a shape derived from the interface
// that marks the method required (same approach as user.ts's account stub views).
type ProfileAccountStub = Required<Pick<GatekeeperUser, "getAuthenticatedProfile">>;
type ProfileHints = Awaited<ReturnType<ProfileAccountStub["getAuthenticatedProfile"]>>;

// Distinguishes "we stopped waiting" from anything the gatekeeper itself could return, null included.
const HINTS_TIMED_OUT = Symbol("profile hints timed out");

/**
 * Ask the gatekeeper for display-name/photo hints, bounded by PROFILE_HINT_TIMEOUT_MS. Never
 * throws: a slow, broken, or hint-less gatekeeper yields `{}` and sign-in carries on.
 */
async function fetchProfileHints(
    account: Fetcher<GatekeeperUser>, log: LoginLogger): Promise<ProfileHints> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    // On timeout the gatekeeper's promise stays pending, but Promise.race has already attached
    // handlers to it, so a late rejection can't surface as an unhandled rejection.
    let hints = await Promise.race([
      (account as unknown as ProfileAccountStub).getAuthenticatedProfile(),
      new Promise<typeof HINTS_TIMED_OUT>(resolve => {
        timer = setTimeout(() => resolve(HINTS_TIMED_OUT), PROFILE_HINT_TIMEOUT_MS);
      }),
    ]);
    if (hints === HINTS_TIMED_OUT) {
      log.warn("gatekeeper profile hints timed out", {
        event: "gatekeeper.login.profile.timeout",
      });
      return {};
    }
    // What came back is another worker's value, only a hints object by declaration; anything else
    // is "no hints", so reading a field off it later can't turn into a sign-in failure.
    if (hints === null || typeof hints !== "object") return {};
    return hints;
  } catch (err) {
    log.warn("gatekeeper profile hints unavailable", {
      event: "gatekeeper.login.profile.failed", error: err,
    });
    return {};
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Store the gatekeeper's photo as the user's avatar, but only when they have none — an avatar the
 * user (or an earlier sign-in) already set is never overwritten. Returns whether it was written.
 * Never throws: a photo the Workshop won't accept, or a KV failure, must not fail a sign-in.
 */
async function seedAvatar(
    env: AvatarEnv, userId: string, photo: NonNullable<ProfileHints["photo"]>,
    log: LoginLogger): Promise<boolean> {
  try {
    if (await userAvatarExists(env, userId)) return false;
    validateAvatarBytes(photo.data);
    await putUserAvatar(env, userId, photo.data);
    return true;
  } catch (err) {
    log.warn("gatekeeper profile photo not seeded", {
      event: "gatekeeper.login.profile.photo.failed", error: err,
    });
    return false;
  }
}

function describeSeeded(name: boolean, photo: boolean): "name+photo" | "name" | "photo" | "none" {
  if (name && photo) return "name+photo";
  if (name) return "name";
  if (photo) return "photo";
  return "none";
}

type PendingResult = { token: string } | { error: string };

/**
 * Bridges a login result from the (separate) OAuth-callback invocation back to the waiting browser.
 *
 * This DO holds no durable storage: a login normally completes within seconds, and the in-flight
 * awaitResult() request keeps the DO alive so the in-memory waiter is reachable when deliver()/fail()
 * fire. If the attempt is abandoned, the client disposes the awaiting RPC (the `attempt` stub) and
 * the DO is simply evicted — no alarm or cleanup needed.
 */
export class PendingLogin extends DurableObject<Cloudflare.Env> {
  // Awaiters from in-flight awaitResult() calls, resolved/rejected when the result arrives.
  #waiters: { resolve: (token: string) => void; reject: (err: Error) => void }[] = [];
  // Stash for the rare case deliver()/fail() arrives before awaitResult() registers a waiter.
  #result?: PendingResult;

  /** Block until the login completes (or fails). */
  async awaitResult(): Promise<string> {
    if (this.#result) {
      const result = this.#result;
      this.#result = undefined;  // one-time use
      if ("token" in result) return result.token;
      throw new Error(result.error);
    }
    return await new Promise<string>((resolve, reject) => {
      this.#waiters.push({ resolve, reject });
    });
  }

  /**
   * Called by LoginConnectCallbackImpl on success: resolve the awaiter (or stash the token if none
   * is waiting yet).
   */
  async deliver(token: string): Promise<void> {
    if (this.#waiters.length > 0) {
      for (const w of this.#waiters) w.resolve(token);
      this.#waiters = [];
    } else {
      this.#result = { token };
    }
  }

  async fail(reason: string): Promise<void> {
    if (this.#waiters.length > 0) {
      for (const w of this.#waiters) w.reject(new Error(reason));
      this.#waiters = [];
    } else {
      this.#result = { error: reason };
    }
  }
}

// `providesProfile` mirrors the vendor's providesAuthProfile declaration, read once when the login
// starts, so complete() knows whether asking for profile hints is worth an RPC.
type LoginCallbackProps = { pendingId: string; vendorId: string; providesProfile: boolean };

export class LoginConnectCallbackImpl
    extends WorkerEntrypoint<Cloudflare.Env, LoginCallbackProps>
    implements GatekeeperConnectCallback {
  #pending() {
    const id = this.ctx.exports.PendingLogin.idFromString(this.ctx.props.pendingId);
    return this.ctx.exports.PendingLogin.get(id);
  }

  async complete(account: Fetcher<GatekeeperUser>, expiresAt?: Date): Promise<void> {
    const loginLogger = logger.with({
      operation: "gatekeeper.login",
      vendorId: this.ctx.props.vendorId,
    });
    const pending = this.#pending();
    // `account` is a call parameter, so Cap'n Web disposes it automatically when this method
    // returns — no explicit disposal needed. We read the verified email to resolve/create the user.
    // The email's local-part seeds the initial display name, like the Cloudflare Access flow, unless
    // the gatekeeper offers a better one (see the profile hints below).
    try {
      const email = await account.getAuthenticatedEmail();
      if (!email) {
        loginLogger.info("gatekeeper login finished", {
          event: "gatekeeper.login.finished", outcome: "no_email",
        });
        await pending.fail("This account has no verified email, so it can't be used to sign in.");
        return;
      }
      // Profile hints (display name, photo) seed the new user's profile. They are gathered and
      // applied before the token is delivered, so the welcome screen sees the seeded state on its
      // first render — and every step of it degrades to "no hints" rather than failing the login.
      const hints = this.ctx.props.providesProfile ? await fetchProfileHints(account, loginLogger)
                                                  : {};
      const nameHint = sanitizeDisplayNameHint(hints.name);
      const userStub = this.ctx.exports.UserDurableObject.get(
          this.ctx.exports.UserDurableObject.idFromName(email));
      // Closed signups block first-time account creation here too (not just password signup); an
      // existing user signing in is unaffected.
      const signupsEnabled = (await readAdminConfig(this.env)).signupsEnabled;
      const login = await userStub.loginOrCreateViaGatekeeper(
          email, signupsEnabled, nameHint ?? undefined);
      if (login === null) {
        loginLogger.info("gatekeeper login finished", {
          event: "gatekeeper.login.finished", outcome: "signups_disabled",
        });
        await pending.fail("New sign-ups are currently disabled on this deployment.");
        return;
      }
      const photoSeeded = hints.photo !== undefined
          && await seedAvatar(this.env, email, hints.photo, loginLogger);
      // For Cloudflare, signing in also links the account for AI Gateway billing: startGatekeeperLogin
      // requested full (non-transient) scopes, so persist the grant as a connected account before
      // handing back the session. Other providers use minimal, transient sign-in grants (no persist).
      if (this.ctx.props.vendorId === CLOUDFLARE_VENDOR_ID) {
        await userStub.linkConnectedAccountFromLogin(account, this.ctx.props.vendorId, expiresAt);
      }
      // Session tokens are "<doName>:<secret>"; PublicApi.authenticate() routes via idFromName of
      // the first part. The user DO is keyed by email, so the prefix must be the email.
      await pending.deliver(`${email}:${login.secret}`);
      // Which hints we managed to seed, never their values: the name and the email are the user's.
      // `nameSeeded` comes from the user DO, so a hint it declined (customised name) reads as none.
      loginLogger.info("gatekeeper login finished", {
        event: "gatekeeper.login.finished", outcome: "ok",
        profileSeeded: describeSeeded(login.nameSeeded, photoSeeded),
      });
    } catch (err) {
      loginLogger.error("gatekeeper login failed", {
        event: "gatekeeper.login.failed", error: err,
      });
      loginLogger.info("gatekeeper login finished", {
        event: "gatekeeper.login.finished", outcome: "error",
      });
      await pending.fail("Sign-in failed. Please try again.");
    }
  }

  /**
   * No-ops: for transient sign-in grants there's nothing persisted to update. For the Cloudflare
   * billing connection (persisted on login) these would ideally flip the account's credential flag,
   * but the callback doesn't carry the user/account identity (it's only learned in complete()). The
   * billing path degrades gracefully regardless — getUsableAccessToken() returns null on expiry and
   * the user falls back to the free tier / a reconnect prompt.
   */
  async credentialsExpired(): Promise<void> {}
  async credentialsRestored(_expiresAt?: Date): Promise<void> {}
}
