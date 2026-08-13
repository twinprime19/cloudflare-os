import { describe, expect, it } from "vitest";

import { grantCoversScopes, normalizeScope } from "../src/microsoft-api";

describe("normalizeScope", () => {
  it("strips the Graph resource prefix and lowercases", () => {
    expect(normalizeScope("https://graph.microsoft.com/Mail.ReadWrite")).toBe("mail.readwrite");
    expect(normalizeScope("  Mail.Readwrite ")).toBe("mail.readwrite");
    expect(normalizeScope("User.Read")).toBe("user.read");
  });

  it("leaves a non-Graph scope alone apart from casing", () => {
    expect(normalizeScope("https://outlook.office.com/Mail.ReadWrite"))
      .toBe("https://outlook.office.com/mail.readwrite");
  });
});

describe("grantCoversScopes", () => {
  it("matches a resource-qualified response scope against a bare request scope", () => {
    expect(grantCoversScopes(
      ["https://graph.microsoft.com/Mail.ReadWrite", "https://graph.microsoft.com/User.Read"],
      ["Mail.ReadWrite"])).toBe(true);
  });

  it("ignores casing differences", () => {
    expect(grantCoversScopes(["mail.readwrite"], ["Mail.ReadWrite"])).toBe(true);
    expect(grantCoversScopes(["Mail.READWRITE"], ["Mail.ReadWrite"])).toBe(true);
  });

  it("treats the reserved OIDC scopes as granted even though Entra omits them", () => {
    expect(grantCoversScopes(
      ["https://graph.microsoft.com/User.Read"],
      ["openid", "profile", "email", "offline_access", "User.Read"])).toBe(true);
  });

  it("reports a genuinely missing permission as missing", () => {
    expect(grantCoversScopes(["https://graph.microsoft.com/User.Read"], ["Mail.ReadWrite"]))
      .toBe(false);
    expect(grantCoversScopes([], ["Mail.ReadWrite"])).toBe(false);
    // A read-only grant does not satisfy a read-write requirement.
    expect(grantCoversScopes(["Mail.Read"], ["Mail.ReadWrite"])).toBe(false);
  });
});
