// Parsing of the SharePoint list URL a user pastes to introduce a list.
//
// This is the only place a browser URL becomes the ids the Graph client works with, so it is strict
// on purpose: the host must really be SharePoint Online, and every path component it hands on must
// survive being split on `/` and re-encoded without changing what it addresses. Everything a page
// URL carries beyond that — the view page, the query string, the fragment — is view state, not
// identity, so it is ignored.
//
// SharePoint's "Copy link" button writes a different shape, `/:l:/r/sites/HR/Lists/Requests`: a
// one-letter type (`l` list, `f` folder, `w` doc …), a mode, then the path. Mode `r` is the only
// one that carries that real path, so it is stripped and parsing continues; the other modes carry
// an opaque share id instead and are refused.
//
// The parser never echoes the URL back in an error. The value came from a user and travels on to a
// toast and to an agent; saying what shape was expected is what helps, and it cannot carry anything
// hostile along with it.

/** SharePoint Online's host suffix. Matched on a label boundary, so `sharepoint.com.evil` fails. */
const SHAREPOINT_HOST_SUFFIX = ".sharepoint.com";

/** The path segment that separates a site's path from a list's name. */
const LISTS_SEGMENT = "lists";

/** A share link's leading type segment, as "Copy link" writes it: `:l:`, `:f:`, `:w:`, `:x:`. */
const SHARE_LINK_TYPE = /^:[a-z]:$/i;

/** The one share mode whose next segments are the real path rather than an opaque share id. */
const SHARE_LINK_PATH_MODE = "r";

/** The shape every error message names, so a user can compare it with what they pasted. */
const EXPECTED_SHAPE =
    "https://<tenant>.sharepoint.com/sites/<Site>/Lists/<List>/AllItems.aspx";

/** A pasted URL that is not a SharePoint list URL this gatekeeper can open. */
export class SharePointUrlError extends Error {
  constructor(problem: string) {
    super(`${problem} Expected a SharePoint list URL like ${EXPECTED_SHAPE}`);
    this.name = "SharePointUrlError";
  }
}

/**
 * Whether a host is SharePoint Online's, matched on a label boundary so `sharepoint.com` itself and
 * `sharepoint.com.evil` both fail.
 *
 * Exported because whoever routes a pasted URL needs the answer before parsing has succeeded: once
 * the host is SharePoint the user meant this resource, so the parser's own refusal — which names the
 * shape to paste — is what should reach them, rather than a message about an unconnectable URL.
 */
export function isSharePointHost(hostname: string): boolean {
  let lowered = hostname.toLowerCase();
  return lowered.endsWith(SHAREPOINT_HOST_SUFFIX) &&
      lowered.length > SHAREPOINT_HOST_SUFFIX.length;
}

/** The parts of a list URL that identify the list, decoded. */
export type SharePointListUrl = {
  /** The SharePoint Online host, lowercased: `contoso.sharepoint.com`. */
  hostname: string;
  /**
   * Server-relative site path, decoded, with no trailing slash: `/sites/HR`, `/teams/HR/sub`, or
   * the empty string for the tenant root site.
   */
  sitePath: string;
  /** The list's URL name, decoded: `DEV Contoso OS Issues Log`. */
  listSegment: string;
};

/** Percent-decode one path component, or throw for an invalid escape sequence. */
function decodeSegment(segment: string): string {
  try {
    return decodeURIComponent(segment);
  } catch {
    throw new SharePointUrlError("That URL contains an invalid percent-escape.");
  }
}

/**
 * Parse a pasted SharePoint list URL into the parts that identify the list.
 *
 * @throws SharePointUrlError when the URL is not an https SharePoint Online list URL.
 */
export function parseSharePointListUrl(url: string): SharePointListUrl {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    throw new SharePointUrlError("That is not a URL.");
  }

  if (parsed.protocol !== "https:") {
    throw new SharePointUrlError("A SharePoint list URL must use https.");
  }

  let hostname = parsed.hostname.toLowerCase();
  if (!isSharePointHost(hostname)) {
    throw new SharePointUrlError(`That URL is not on a ${SHAREPOINT_HOST_SUFFIX} host.`);
  }

  // `pathname` is still percent-encoded here, which is what makes splitting on `/` correct: a `%2F`
  // inside a component cannot be mistaken for a separator. Components are decoded one at a time
  // below, after the split.
  let raw = parsed.pathname.split("/").filter(segment => segment !== "");

  // A "Copy link" URL puts `/:<type>:/<mode>/` in front of the path the address bar would show.
  // Only mode `r` is followed by that path; the others name the share, not the list, so nothing
  // here could resolve them.
  if (raw.length > 0 && SHARE_LINK_TYPE.test(raw[0])) {
    if ((raw[1] ?? "").toLowerCase() !== SHARE_LINK_PATH_MODE) {
      throw new SharePointUrlError(
          "That sharing link identifies a share rather than the list's path. Open the list and " +
          "copy the URL from your browser's address bar instead.");
    }
    raw = raw.slice(2);
  }

  let listsAt = raw.findIndex(segment => segment.toLowerCase() === LISTS_SEGMENT);
  if (listsAt === -1) {
    throw new SharePointUrlError("That URL has no /Lists/ segment, so it does not name a list.");
  }

  let listSegment = decodeSegment(raw[listsAt + 1] ?? "");
  if (!listSegment) {
    throw new SharePointUrlError("That URL has no list name after /Lists/.");
  }

  let sitePath = raw.slice(0, listsAt).map(segment => {
    let decoded = decodeSegment(segment);
    // A decoded `/` or `\` would turn one component into two the next time this path is split.
    // `.` and `..` are already collapsed by URL parsing before they get here (in both their literal
    // and percent-encoded forms), so refusing them is belt-and-braces; neither occurs in a real site
    // path either way.
    if (decoded === "." || decoded === ".." || decoded.includes("/") || decoded.includes("\\")) {
      throw new SharePointUrlError("That URL has a site path this cannot address safely.");
    }
    return decoded;
  }).join("/");

  return {
    hostname,
    sitePath: sitePath ? `/${sitePath}` : "",
    listSegment,
  };
}
