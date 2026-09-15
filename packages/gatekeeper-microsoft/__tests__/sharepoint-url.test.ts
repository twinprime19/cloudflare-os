import { describe, expect, it } from "vitest";

import { SharePointUrlError, parseSharePointListUrl } from "../src/sharepoint-url";

/** Every shape a user can realistically paste, and what the list's identity is inside it. */
const ACCEPTED: { name: string; url: string; hostname: string; sitePath: string; list: string }[] = [
  {
    name: "a site list's default view",
    url: "https://contoso.sharepoint.com/sites/HR/Lists/Requests/AllItems.aspx",
    hostname: "contoso.sharepoint.com",
    sitePath: "/sites/HR",
    list: "Requests",
  },
  {
    name: "a team site with no view page",
    url: "https://contoso.sharepoint.com/teams/HR/Lists/Requests",
    hostname: "contoso.sharepoint.com",
    sitePath: "/teams/HR",
    list: "Requests",
  },
  {
    name: "the tenant root site",
    url: "https://contoso.sharepoint.com/Lists/Requests/",
    hostname: "contoso.sharepoint.com",
    sitePath: "",
    list: "Requests",
  },
  {
    name: "a subsite, with the list name percent-encoded",
    url: "https://contoso.sharepoint.com/sites/HR/sub/Lists/Req%20uests/",
    hostname: "contoso.sharepoint.com",
    sitePath: "/sites/HR/sub",
    list: "Req uests",
  },
  {
    // The URL the owner actually pasted, from the list this feature was built for.
    name: "a real list URL with a view page and a query string",
    url: "https://contoso.sharepoint.com/sites/Engineering/Lists/" +
        "DEV%20Contoso%20OS%20Issues%20Log/AllItems.aspx?npsAction=createList",
    hostname: "contoso.sharepoint.com",
    sitePath: "/sites/Engineering",
    list: "DEV Contoso OS Issues Log",
  },
  {
    name: "a lowercased /lists/ segment",
    url: "https://contoso.sharepoint.com/sites/HR/lists/Requests/AllItems.aspx",
    hostname: "contoso.sharepoint.com",
    sitePath: "/sites/HR",
    list: "Requests",
  },
  {
    // The share link SharePoint's "Copy link" button produced for that same list.
    name: "a share link for that list, with the /:l:/r/ prefix its Copy link button adds",
    url: "https://contoso.sharepoint.com/:l:/r/sites/Engineering/Lists/" +
        "DEV%20Contoso%20OS%20Issues%20Log?e=AbC123",
    hostname: "contoso.sharepoint.com",
    sitePath: "/sites/Engineering",
    list: "DEV Contoso OS Issues Log",
  },
  {
    name: "a share link for a root-site list",
    url: "https://contoso.sharepoint.com/:l:/r/Lists/Requests",
    hostname: "contoso.sharepoint.com",
    sitePath: "",
    list: "Requests",
  },
  {
    name: "a share link for a team site's list",
    url: "https://contoso.sharepoint.com/:l:/r/teams/HR/Lists/Requests/AllItems.aspx",
    hostname: "contoso.sharepoint.com",
    sitePath: "/teams/HR",
    list: "Requests",
  },
  {
    name: "a share link whose type and mode are upper case",
    url: "https://contoso.sharepoint.com/:L:/R/sites/HR/Lists/Requests",
    hostname: "contoso.sharepoint.com",
    sitePath: "/sites/HR",
    list: "Requests",
  },
  {
    name: "a host in mixed case, with a fragment",
    url: "https://Contoso.SharePoint.com/sites/HR/Lists/Requests/AllItems.aspx#view",
    hostname: "contoso.sharepoint.com",
    sitePath: "/sites/HR",
    list: "Requests",
  },
];

/** Everything that must not be accepted, and the part of the message that says why. */
const REJECTED: { name: string; url: string; because: RegExp }[] = [
  {
    name: "a host that is not SharePoint Online",
    url: "https://contoso.example.com/sites/HR/Lists/Requests/AllItems.aspx",
    because: /not on a \.sharepoint\.com host/,
  },
  {
    name: "a host that only ends in the suffix's text",
    url: "https://sharepoint.com.evil.example/sites/HR/Lists/Requests/AllItems.aspx",
    because: /not on a \.sharepoint\.com host/,
  },
  {
    name: "the bare suffix as a host",
    url: "https://sharepoint.com/sites/HR/Lists/Requests",
    because: /not on a \.sharepoint\.com host/,
  },
  {
    name: "a document library, which has no /Lists/ segment",
    url: "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/Forms/AllItems.aspx",
    because: /no \/Lists\/ segment/,
  },
  {
    name: "a /Lists/ segment with nothing after it",
    url: "https://contoso.sharepoint.com/sites/HR/Lists/",
    because: /no list name after \/Lists\//,
  },
  {
    name: "a share link in a mode that carries a share id instead of the list's path",
    url: "https://contoso.sharepoint.com/:l:/s/sites/HR/EQ1Ab2Cd3Ef4",
    because: /copy the URL from your browser's address bar/,
  },
  {
    name: "a guest-mode share link, even when a /Lists/ segment follows it",
    url: "https://contoso.sharepoint.com/:l:/g/sites/HR/Lists/Requests",
    because: /copy the URL from your browser's address bar/,
  },
  {
    name: "a share link with a type but nothing after it",
    url: "https://contoso.sharepoint.com/:l:/",
    because: /copy the URL from your browser's address bar/,
  },
  {
    name: "plain http, which would send the token in the clear if it were ever followed",
    url: "http://contoso.sharepoint.com/sites/HR/Lists/Requests",
    because: /must use https/,
  },
  {
    name: "something that is not a URL at all",
    url: "contoso.sharepoint.com/sites/HR/Lists/Requests",
    because: /not a URL/,
  },
  {
    name: "a site path component that decodes to a separator",
    url: "https://contoso.sharepoint.com/sites/HR%2Fsub/Lists/Requests",
    because: /cannot address safely/,
  },
  {
    name: "an invalid percent-escape",
    url: "https://contoso.sharepoint.com/sites/HR/Lists/Req%zzuests",
    because: /invalid percent-escape/,
  },
];

describe("parseSharePointListUrl", () => {
  for (const testCase of ACCEPTED) {
    it(`parses ${testCase.name}`, () => {
      expect(parseSharePointListUrl(testCase.url)).toEqual({
        hostname: testCase.hostname,
        sitePath: testCase.sitePath,
        listSegment: testCase.list,
      });
    });
  }

  for (const testCase of REJECTED) {
    it(`rejects ${testCase.name}`, () => {
      expect(() => parseSharePointListUrl(testCase.url)).toThrow(SharePointUrlError);
      expect(() => parseSharePointListUrl(testCase.url)).toThrow(testCase.because);
    });
  }

  it("names the expected shape in every rejection, so a user can compare it with what they pasted",
     () => {
       for (const testCase of REJECTED) {
         expect(() => parseSharePointListUrl(testCase.url))
           .toThrow(/https:\/\/<tenant>\.sharepoint\.com\/sites\/<Site>\/Lists\/<List>/);
       }
     });

  it("cannot carry a relative site path, because URL parsing collapses dot segments first", () => {
    // Both the literal and the percent-encoded form are removed by the URL parser, so a traversal
    // resolves to a shorter real path rather than reaching the site-path guard at all.
    expect(parseSharePointListUrl("https://contoso.sharepoint.com/sites/%2E%2E/Lists/Requests"))
      .toEqual({ hostname: "contoso.sharepoint.com", sitePath: "", listSegment: "Requests" });
    expect(parseSharePointListUrl("https://contoso.sharepoint.com/sites/../Lists/Requests"))
      .toEqual({ hostname: "contoso.sharepoint.com", sitePath: "", listSegment: "Requests" });
  });

  it("never echoes the pasted URL back, since it travels on to a toast and to an agent", () => {
    const hostile = "https://contoso.example.com/sites/HR/Lists/<script>alert(1)</script>";

    expect(() => parseSharePointListUrl(hostile)).toThrow(SharePointUrlError);
    try {
      parseSharePointListUrl(hostile);
    } catch (error) {
      expect((error as Error).message).not.toContain("script");
    }
  });
});
