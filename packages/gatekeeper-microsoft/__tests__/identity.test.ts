import { describe, expect, it } from "vitest";

import {
  classifyTokenFailure, decodeIdTokenClaims, resolveVerifiedEmail, type GraphUserProfile,
} from "../src/microsoft-api";

const TENANT = "11111111-2222-3333-4444-555555555555";
const VERIFIED_DOMAINS = ["contoso.com", "CONTOSO.onmicrosoft.com"];
const CLAIMS = { tid: TENANT, oid: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" };

function member(overrides: Partial<GraphUserProfile> = {}): GraphUserProfile {
  return {
    id: "user-1",
    displayName: "Alex Morgan",
    userType: "Member",
    userPrincipalName: "alex@contoso.com",
    mail: "alex@contoso.com",
    ...overrides,
  };
}

describe("resolveVerifiedEmail", () => {
  it("accepts a tenant member whose mail sits on a verified domain", () => {
    expect(resolveVerifiedEmail(member(), CLAIMS, VERIFIED_DOMAINS, TENANT))
      .toBe("alex@contoso.com");
  });

  it("lowercases the address so one identity cannot become two accounts", () => {
    const profile = member({ mail: undefined, userPrincipalName: "Alex.NG@Contoso.COM" });
    expect(resolveVerifiedEmail(profile, CLAIMS, VERIFIED_DOMAINS, TENANT))
      .toBe("alex.ng@contoso.com");
  });

  it("matches verified domains case-insensitively", () => {
    const profile = member({
      mail: "svc@contoso.onmicrosoft.com", userPrincipalName: "svc@contoso.onmicrosoft.com",
    });
    expect(resolveVerifiedEmail(profile, CLAIMS, VERIFIED_DOMAINS, TENANT))
      .toBe("svc@contoso.onmicrosoft.com");
  });

  it("rejects a guest whose invited external address is populated in mail", () => {
    // The takeover case: a B2B guest is a real principal of this tenant, so the tid check passes,
    // and `mail` carries an address at a domain this tenant does not own.
    const guest = member({
      userType: "Guest",
      mail: "attacker@gmail.com",
      userPrincipalName: "attacker_gmail.com#EXT#@contoso.onmicrosoft.com",
    });
    expect(resolveVerifiedEmail(guest, CLAIMS, VERIFIED_DOMAINS, TENANT)).toBeNull();
  });

  it("rejects a profile with no userType, since a guest is then indistinguishable", () => {
    const profile = member({ userType: undefined });
    expect(resolveVerifiedEmail(profile, CLAIMS, VERIFIED_DOMAINS, TENANT)).toBeNull();
  });

  it("rejects an #EXT# principal name even when it claims to be a member with on-domain mail", () => {
    const profile = member({
      userType: "Member",
      mail: "alex@contoso.com",
      userPrincipalName: "attacker_gmail.com#EXT#@contoso.onmicrosoft.com",
    });
    expect(resolveVerifiedEmail(profile, CLAIMS, VERIFIED_DOMAINS, TENANT)).toBeNull();
  });

  it("rejects mail on a domain the tenant has not verified", () => {
    const profile = member({ mail: "alex@personal.example" });
    expect(resolveVerifiedEmail(profile, CLAIMS, VERIFIED_DOMAINS, TENANT)).toBeNull();
  });

  it("rejects a token minted for a different tenant", () => {
    const otherTenant = { ...CLAIMS, tid: "99999999-2222-3333-4444-555555555555" };
    expect(resolveVerifiedEmail(member(), otherTenant, VERIFIED_DOMAINS, TENANT)).toBeNull();
  });

  it("rejects when no id_token claims were recorded", () => {
    expect(resolveVerifiedEmail(member(), undefined, VERIFIED_DOMAINS, TENANT)).toBeNull();
  });

  it("rejects when the deployment names no tenant", () => {
    expect(resolveVerifiedEmail(member(), CLAIMS, VERIFIED_DOMAINS, "")).toBeNull();
  });

  it("rejects when the tenant reports no verified domains", () => {
    expect(resolveVerifiedEmail(member(), CLAIMS, [], TENANT)).toBeNull();
  });

  it("rejects malformed addresses", () => {
    for (const mail of ["alex", "alex@", "@contoso.com", "a@b@contoso.com",
                        "al ex@contoso.com", ""]) {
      const profile = member({ mail, userPrincipalName: mail });
      expect(resolveVerifiedEmail(profile, CLAIMS, VERIFIED_DOMAINS, TENANT)).toBeNull();
    }
  });

  it("falls back to the principal name when mail is unset", () => {
    const profile = member({ mail: null });
    expect(resolveVerifiedEmail(profile, CLAIMS, VERIFIED_DOMAINS, TENANT))
      .toBe("alex@contoso.com");
  });

  it("rejects a missing profile", () => {
    expect(resolveVerifiedEmail(undefined, CLAIMS, VERIFIED_DOMAINS, TENANT)).toBeNull();
  });
});

describe("classifyTokenFailure", () => {
  const permanentCodes = [
    { code: 700082, what: "refresh token expired through inactivity" },
    { code: 50173, what: "grant invalidated by a password change or admin revocation" },
    { code: 7000222, what: "client secret expired" },
    { code: 65001, what: "consent not granted" },
  ];

  for (const { code, what } of permanentCodes) {
    it(`treats AADSTS${code} (${what}) as permanent`, () => {
      const failure = classifyTokenFailure({ error: "invalid_grant", error_codes: [code] });
      expect(failure.permanent).toBe(true);
      expect(failure.codes).toContain(code);
    });
  }

  it("treats the interaction-required family as permanent, whatever code it carries", () => {
    // Conditional Access step-up and MFA challenges surface here with codes we deliberately do not
    // enumerate; without this the account refreshes forever and never prompts a reconnect.
    for (const error of ["interaction_required", "login_required", "consent_required"]) {
      const failure = classifyTokenFailure({
        error,
        error_codes: [50076],
        error_description: "AADSTS50076: Due to a configuration change made by your administrator…",
      });
      expect(failure.permanent).toBe(true);
      expect(failure.message).toMatch(/reconnect the account/i);
    }
  });

  it("classifies interaction_required even when the body carries no codes at all", () => {
    expect(classifyTokenFailure({ error: "interaction_required" }).permanent).toBe(true);
  });

  it("treats an uncatalogued code as transient so a live account is never killed by guesswork", () => {
    const failure = classifyTokenFailure({ error: "temporarily_unavailable", error_codes: [90033] });
    expect(failure.permanent).toBe(false);
    expect(failure.message).toMatch(/try again/i);
  });

  it("treats an unparseable body as transient", () => {
    expect(classifyTokenFailure(undefined).permanent).toBe(false);
    expect(classifyTokenFailure({}).permanent).toBe(false);
  });

  it("reads the code out of the description when no code array is present", () => {
    const failure = classifyTokenFailure({
      error: "invalid_grant",
      error_description: "AADSTS700082: The refresh token has expired due to inactivity.",
    });
    expect(failure.permanent).toBe(true);
    expect(failure.codes).toContain(700082);
  });

  it("names an administrator, not the user, when the client secret is the problem", () => {
    const failure = classifyTokenFailure({ error_codes: [7000222] });
    expect(failure.permanent).toBe(true);
    expect(failure.message).toMatch(/administrator must rotate/i);
  });

  it("caps provider text so an oversized description cannot become the whole error", () => {
    const failure = classifyTokenFailure({ error_description: "x".repeat(5000) });
    expect(failure.message.length).toBeLessThan(700);
  });
});

function base64url(value: unknown): string {
  return Buffer.from(JSON.stringify(value), "utf8").toString("base64url");
}

function idToken(payload: unknown): string {
  return `${base64url({ alg: "RS256" })}.${base64url(payload)}.signature`;
}

describe("decodeIdTokenClaims", () => {
  it("reads tid and oid", () => {
    expect(decodeIdTokenClaims(idToken({ tid: "tenant", oid: "object", sub: "s" })))
      .toEqual({ tid: "tenant", oid: "object" });
  });

  it("ignores non-string claims", () => {
    expect(decodeIdTokenClaims(idToken({ tid: 7, oid: null }))).toEqual({
      tid: undefined, oid: undefined,
    });
  });

  it("returns undefined for anything that is not a three-part token", () => {
    expect(decodeIdTokenClaims(undefined)).toBeUndefined();
    expect(decodeIdTokenClaims("")).toBeUndefined();
    expect(decodeIdTokenClaims("not.a.jwt")).toBeUndefined();
    expect(decodeIdTokenClaims("only.two")).toBeUndefined();
  });
});
