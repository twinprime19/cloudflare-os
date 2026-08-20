import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  SCAN_ROOT,
  findStringCandidates,
  isAllowlisted,
  isExcludedFile,
  isIdentifierLike,
  isLocalizableAttribute,
  isLocalizableAttributeValue,
  isLocalizedContainer,
  isLoneLabelWord,
  isPathLike,
  isProperNoun,
  isTechnicalLiteralElement,
  isUrlLike,
  isUserFacingText,
  listSourceFiles,
  normalizeText,
  parseAllowlist,
  type StringCandidate,
} from "./detect.ts";

function candidates(source: string): StringCandidate[] {
  return findStringCandidates("packages/workshop-frontend/src/Sample.tsx", source);
}

function texts(source: string): string[] {
  return candidates(source).map((candidate) => candidate.text);
}

describe("normalizeText", () => {
  it("collapses the indentation JSX leaves inside a text node", () => {
    assert.equal(normalizeText("\n      Save this\n      workspace\n    "), "Save this workspace");
  });
});

describe("isUserFacingText", () => {
  it("accepts a phrase with letters and a space", () => {
    assert.equal(isUserFacingText("Delete this workspace"), true);
  });

  it("accepts a phrase whose letters are not Latin", () => {
    assert.equal(isUserFacingText("Xoá không gian"), true);
  });

  it("accepts a phrase that starts lowercase", () => {
    assert.equal(isUserFacingText("no gadgets yet"), true);
  });

  it("accepts a lone capitalized word, which reads as a label", () => {
    assert.equal(isUserFacingText("Cancel"), true);
    assert.equal(isUserFacingText("Loading…"), true);
  });

  it("rejects a lone lowercase word, which is usually half of an interpolated sentence", () => {
    assert.equal(isUserFacingText("characters"), false);
    assert.equal(isUserFacingText("of"), false);
  });

  it("rejects a lone brand name", () => {
    assert.equal(isUserFacingText("Slack"), false);
  });

  it("rejects a string with no letters at all", () => {
    assert.equal(isUserFacingText("— 42 · 7%"), false);
  });

  it("rejects whitespace and empty strings", () => {
    assert.equal(isUserFacingText("   \n  "), false);
    assert.equal(isUserFacingText(""), false);
  });
});

describe("isUrlLike", () => {
  it("rejects a bare URL", () => {
    assert.equal(isUrlLike("https://developers.cloudflare.com/workers"), true);
    assert.equal(isUserFacingText("https://developers.cloudflare.com/workers"), false);
  });

  it("rejects protocol-relative and non-http schemes", () => {
    assert.equal(isUrlLike("//cdn.example.com/logo.svg"), true);
    assert.equal(isUrlLike("mailto:support@example.com"), true);
  });

  it("accepts prose that merely mentions a link", () => {
    assert.equal(isUrlLike("See https://example.com for details"), false);
    assert.equal(isUserFacingText("See https://example.com for details"), true);
  });
});

describe("isPathLike", () => {
  it("rejects absolute, relative, and multi-segment paths", () => {
    assert.equal(isPathLike("/workspaces/new"), true);
    assert.equal(isPathLike("./components/Button"), true);
    assert.equal(isPathLike("src/components/Button.tsx"), true);
  });

  it("rejects a bare filename", () => {
    assert.equal(isPathLike("logo.svg"), true);
    assert.equal(isUserFacingText("logo.svg"), false);
  });

  it("accepts a sentence that ends in a full stop", () => {
    assert.equal(isPathLike("All done."), false);
    assert.equal(isUserFacingText("All done."), true);
  });
});

describe("isIdentifierLike", () => {
  it("rejects camelCase, PascalCase with two capitals, snake_case, kebab-case and SCREAMING", () => {
    assert.equal(isIdentifierLike("onClickHandler"), true);
    assert.equal(isIdentifierLike("GitHub"), true);
    assert.equal(isIdentifierLike("auto_approve"), true);
    assert.equal(isIdentifierLike("aria-label"), true);
    assert.equal(isIdentifierLike("MAX_TOKENS"), true);
  });

  it("rejects a dotted key", () => {
    assert.equal(isIdentifierLike("workshop.react-root"), true);
  });

  it("accepts a single capitalized word, which is a label rather than an identifier", () => {
    assert.equal(isIdentifierLike("Close"), false);
  });

  it("accepts anything containing a space", () => {
    assert.equal(isIdentifierLike("onClick fired"), false);
  });
});

describe("isLoneLabelWord", () => {
  it("accepts a capitalized word of four letters or more", () => {
    assert.equal(isLoneLabelWord("Cancel"), true);
    assert.equal(isLoneLabelWord("Save"), true);
  });

  it("accepts the trailing punctuation UI copy uses", () => {
    assert.equal(isLoneLabelWord("Loading…"), true);
    assert.equal(isLoneLabelWord("Authenticating..."), true);
    assert.equal(isLoneLabelWord("Ready?"), true);
  });

  it("rejects a trailing colon, which a URL scheme is indistinguishable from", () => {
    assert.equal(isLoneLabelWord("Note:"), false);
  });

  it("rejects a word shorter than four letters", () => {
    assert.equal(isLoneLabelWord("Add"), false);
    assert.equal(isLoneLabelWord("Esc"), false);
  });

  it("rejects a lowercase word", () => {
    assert.equal(isLoneLabelWord("selected"), false);
  });

  it("rejects anything with a space", () => {
    assert.equal(isLoneLabelWord("Save changes"), false);
  });

  it("rejects brands, identifiers, paths and URLs", () => {
    assert.equal(isLoneLabelWord("Notion"), false);
    assert.equal(isLoneLabelWord("GitHub"), false);
    assert.equal(isLoneLabelWord("onClick"), false);
    assert.equal(isLoneLabelWord("logo.svg"), false);
    assert.equal(isLoneLabelWord("https://example.com"), false);
  });

  it("accepts a hyphen, which joins letters into one word rather than making two", () => {
    assert.equal(isLoneLabelWord("Re-authenticating…"), true);
    assert.equal(isLoneLabelWord("Sign-in"), true);
    assert.equal(isLoneLabelWord("Auto-approve"), true);
  });

  it("accepts an apostrophe, typographic or ASCII", () => {
    assert.equal(isLoneLabelWord("Don’t"), true);
    assert.equal(isLoneLabelWord("Don't"), true);
  });

  it("counts four letters across the whole word, not per segment", () => {
    // "Re" alone would never qualify; the word it is part of does.
    assert.equal(isLoneLabelWord("Re-run"), true);
    assert.equal(isLoneLabelWord("A-ok"), false);
  });

  it("rejects a hyphenated identifier, which capitalizes every segment", () => {
    assert.equal(isLoneLabelWord("Content-Type"), false);
    assert.equal(isLoneLabelWord("X-Frame-Options"), false);
    assert.equal(isLoneLabelWord("kebab-case"), false);
  });

  it("agrees with isUserFacingText, which shares the rule", () => {
    for (const word of ["Re-authenticating…", "Don’t", "Cancel"]) {
      assert.equal(isUserFacingText(word), true, word);
    }
    for (const word of ["Content-Type", "A-ok", "kebab-case"]) {
      assert.equal(isUserFacingText(word), false, word);
    }
  });
});

describe("isProperNoun", () => {
  it("knows the brands this repository integrates with and renders", () => {
    assert.equal(isProperNoun("Slack"), true);
    assert.equal(isProperNoun("Cloudflare"), true);
    assert.equal(isProperNoun("Home Assistant"), true);
  });

  it("ignores trailing UI punctuation", () => {
    assert.equal(isProperNoun("Cloudflare…"), true);
  });

  it("leaves copy that merely mentions a brand alone", () => {
    assert.equal(isProperNoun("Add credits in Cloudflare"), false);
    assert.equal(isUserFacingText("Add credits in Cloudflare"), true);
  });

  it("does not claim generic words that happen to name a package", () => {
    for (const word of ["Context", "Email", "Scheduler", "Portal"]) {
      assert.equal(isProperNoun(word), false, word);
    }
  });
});

describe("isTechnicalLiteralElement", () => {
  it("knows the elements whose text is a literal rather than prose", () => {
    for (const tag of ["code", "kbd", "pre", "samp", "var"]) {
      assert.equal(isTechnicalLiteralElement(tag), true, tag);
    }
    assert.equal(isTechnicalLiteralElement("p"), false);
  });
});

describe("isLocalizableAttribute", () => {
  it("covers exactly the attributes that carry copy", () => {
    for (const name of ["alt", "aria-description", "aria-label", "placeholder", "title"]) {
      assert.equal(isLocalizableAttribute(name), true, name);
    }
  });

  it("excludes attributes that never carry copy", () => {
    for (const name of ["className", "data-testid", "id", "key", "href", "type", "role"]) {
      assert.equal(isLocalizableAttribute(name), false, name);
    }
  });
});

describe("isLocalizableAttributeValue", () => {
  it("accepts a lone capitalized word of four letters or more", () => {
    assert.equal(isLocalizableAttributeValue("aria-label", "Close"), true);
    assert.equal(isLocalizableAttributeValue("placeholder", "Search"), true);
  });

  it("rejects a capitalized word shorter than four letters", () => {
    assert.equal(isLocalizableAttributeValue("aria-label", "Add"), false);
  });

  it("accepts a phrase", () => {
    assert.equal(isLocalizableAttributeValue("placeholder", "Ask anything…"), true);
  });

  it("rejects the same value on an attribute that is not whitelisted", () => {
    assert.equal(isLocalizableAttributeValue("className", "Ask anything…"), false);
    assert.equal(isLocalizableAttributeValue("data-label", "Close"), false);
  });

  it("rejects identifiers, paths, URLs and value-less strings", () => {
    assert.equal(isLocalizableAttributeValue("title", "GitHub"), false);
    assert.equal(isLocalizableAttributeValue("alt", "logo.svg"), false);
    assert.equal(isLocalizableAttributeValue("title", "https://example.com"), false);
    assert.equal(isLocalizableAttributeValue("alt", "—"), false);
  });
});

describe("isLocalizedContainer", () => {
  it("knows the Lingui elements that localize their own children", () => {
    assert.equal(isLocalizedContainer("Trans"), true);
    assert.equal(isLocalizedContainer("Plural"), true);
    assert.equal(isLocalizedContainer("div"), false);
  });
});

describe("isExcludedFile", () => {
  it("excludes tests, generated code, dependencies and build output", () => {
    assert.equal(isExcludedFile("packages/workshop-frontend/src/ShareModal.test.tsx"), true);
    assert.equal(isExcludedFile("packages/workshop-frontend/src/routeTree.gen.ts"), true);
    assert.equal(isExcludedFile("packages/workshop-frontend/src/generated/app.tsx"), true);
    assert.equal(isExcludedFile("node_modules/thing/index.tsx"), true);
    assert.equal(isExcludedFile("packages/workshop-frontend/dist/assets/index.js"), true);
  });

  it("includes ordinary components, on either path separator", () => {
    assert.equal(isExcludedFile("packages/workshop-frontend/src/ShareModal.tsx"), false);
    assert.equal(isExcludedFile("packages\\workshop-frontend\\src\\ShareModal.tsx"), false);
  });
});

describe("findStringCandidates: JSX text", () => {
  it("finds a phrase in element text", () => {
    assert.deepEqual(texts("<p>Delete this workspace?</p>"), ["Delete this workspace?"]);
  });

  it("reports the text with indentation collapsed but keeps the source span verbatim", () => {
    const [candidate] = candidates("<p>\n  Delete this\n  workspace?\n</p>");
    assert.equal(candidate.text, "Delete this workspace?");
    assert.equal(candidate.raw, "Delete this\n  workspace?");
    assert.equal(candidate.kind, "jsx-text");
    assert.equal(candidate.parentTag, "p");
    assert.equal(candidate.line, 2);
  });

  it("skips text already inside a Lingui element, however deeply nested", () => {
    assert.deepEqual(texts("<Trans>Delete this workspace?</Trans>"), []);
    assert.deepEqual(texts("<Trans>Delete <b>this workspace</b> now?</Trans>"), []);
  });

  it("skips whitespace between elements", () => {
    assert.deepEqual(texts("<div>\n  <Icon />\n  <Icon />\n</div>"), []);
  });

  it("finds a lone capitalized word, and skips a run with no letters", () => {
    assert.deepEqual(texts("<button>Cancel</button>"), ["Cancel"]);
    assert.deepEqual(texts("<span>· 42 ·</span>"), []);
  });

  it("skips a lone lowercase word and a lone brand", () => {
    assert.deepEqual(texts("<span>selected</span>"), []);
    assert.deepEqual(texts("<span>Notion</span>"), []);
  });

  it("skips everything inside a technical-literal element, however deeply nested", () => {
    assert.deepEqual(texts("<kbd>Enter</kbd>"), []);
    assert.deepEqual(texts("<pre>Could not reach the server</pre>"), []);
    assert.deepEqual(texts("<code><span>Cancel</span></code>"), []);
  });

  it("still finds an attribute on an element inside a technical-literal element", () => {
    assert.deepEqual(texts('<pre><span title="Copy the command">npm i</span></pre>'),
      ["Copy the command"]);
  });

  it("flags each side of an interpolation separately and says so", () => {
    const found = candidates("<p>Always approve {label} for this gadget?</p>");
    assert.deepEqual(found.map((c) => c.text), ["Always approve", "for this gadget?"]);
    assert.equal(found.every((c) => c.interpolated), true);
  });

  it("marks text with no sibling expression as not interpolated", () => {
    const [candidate] = candidates("<p>Always approve this gadget?</p>");
    assert.equal(candidate.interpolated, false);
  });

  it("reports a span the caller can slice out of the source unchanged", () => {
    const source = "<p>\n  Delete this workspace?\n</p>";
    const [candidate] = candidates(source);
    assert.equal(source.slice(candidate.textStart, candidate.textEnd), candidate.raw);
  });
});

// Offsets are UTF-16 code units in both the parser and `String.prototype.slice`, and an HTML entity
// is just a run of ordinary characters to a JSX text node. Neither is obvious, and a wrap that
// re-slices the source depends on both, so they are pinned here rather than assumed.
describe("findStringCandidates: entities and non-ASCII text", () => {
  it("treats an HTML entity as part of the text and keeps the source spelling", () => {
    const [candidate] = candidates("<p>Tom &amp; Jerry drink tea</p>");
    assert.equal(candidate.text, "Tom &amp; Jerry drink tea");
    assert.equal(candidate.raw, "Tom &amp; Jerry drink tea");
  });

  it("does not let a non-breaking-space entity stand in for a real space", () => {
    // `&nbsp;` is six ordinary characters in the source, so it never satisfies the space rule; the
    // first case only has a lone word to offer and the lone-label-word rule rejects the joined form.
    assert.deepEqual(texts("<span>Save&nbsp;now</span>"), []);
    assert.deepEqual(texts("<span>Loading&nbsp;the workspace</span>"),
      ["Loading&nbsp;the workspace"]);
  });

  it("keeps a curly-quote entity inside the span rather than beside it", () => {
    const source = "<p>Applies on each user&rsquo;s next connection.</p>";
    const [candidate] = candidates(source);
    assert.equal(source.slice(candidate.textStart, candidate.textEnd), candidate.raw);
    assert.equal(candidate.raw, "Applies on each user&rsquo;s next connection.");
  });

  it("reports a span that slices back exactly across an emoji", () => {
    // An emoji is a surrogate pair: two UTF-16 units, one code point. An offset that counted code
    // points would cut it in half here.
    const source = "<p>🎉 New blueprints are available</p>";
    const [candidate] = candidates(source);
    assert.equal(source.slice(candidate.textStart, candidate.textEnd), candidate.raw);
    assert.equal(candidate.raw, "🎉 New blueprints are available");
  });

  it("reports an attribute span that slices back exactly across an emoji", () => {
    const source = '<input placeholder="e.g. 🎉 imports are supported now" />';
    const [candidate] = candidates(source);
    assert.equal(source.slice(candidate.textStart, candidate.textEnd), candidate.raw);
    assert.equal(candidate.raw, "e.g. 🎉 imports are supported now");
  });

  it("finds copy written in a non-Latin script", () => {
    const source = "<p>Xoá không gian làm việc này?</p>";
    const [candidate] = candidates(source);
    assert.equal(source.slice(candidate.textStart, candidate.textEnd), candidate.raw);
    assert.equal(candidate.text, "Xoá không gian làm việc này?");
  });

  it("skips a run of emoji with no letters in it", () => {
    assert.deepEqual(texts("<span>🎉 🎊 ✨</span>"), []);
  });
});

describe("findStringCandidates: attributes", () => {
  it("finds a whitelisted attribute's literal value", () => {
    const [candidate] = candidates('<input placeholder="Ask anything" />');
    assert.equal(candidate.kind, "jsx-attribute");
    assert.equal(candidate.attribute, "placeholder");
    assert.equal(candidate.text, "Ask anything");
  });

  it("finds a value written as a braced literal", () => {
    assert.deepEqual(texts('<button aria-label={"Close"} />'), ["Close"]);
  });

  it("skips attributes that are not whitelisted, including data attributes", () => {
    assert.deepEqual(texts('<div className="flex items-center" data-label="Close me" id="a b" />'), []);
  });

  it("skips a value that is an expression rather than a literal", () => {
    assert.deepEqual(texts("<button aria-label={t`Close`} />"), []);
    assert.deepEqual(texts("<button aria-label={label} />"), []);
  });

  // A template value yields no `jsx-attribute` candidate, because this collector takes quoted
  // strings only -- but it is still copy, so the expression collector reports it instead. Asserting
  // it is skipped outright is what kept it invisible to both the codemod and the guard.
  it("leaves a template value to the expression collector rather than dropping it", () => {
    const found = candidates("<button aria-label={`Close ${name}`} />");
    assert.deepEqual(found.map((c) => c.kind), ["jsx-expression"]);
    assert.deepEqual(found.map((c) => c.text), ["Close {0}"]);
  });

  it("reports a span that excludes the quotes", () => {
    const source = '<input placeholder="Ask anything" />';
    const [candidate] = candidates(source);
    assert.equal(source.slice(candidate.textStart, candidate.textEnd), "Ask anything");
    assert.equal(source.slice(candidate.start, candidate.end), '"Ask anything"');
  });

  it("still finds an attribute on an element inside a Lingui element", () => {
    // `<Trans>` localizes its children, not the attributes of the markup within it.
    assert.deepEqual(texts('<Trans>Read <a title="Open the docs">the docs</a></Trans>'),
      ["Open the docs"]);
  });
});

describe("parseAllowlist and isAllowlisted", () => {
  const candidate: StringCandidate = {
    kind: "jsx-text",
    file: "packages/workshop-frontend/src/Api.tsx",
    text: "GET the thing",
    raw: "GET the thing",
    start: 0, end: 13, textStart: 0, textEnd: 13, line: 1, column: 1,
  };

  it("ignores comments and blank lines", () => {
    const allowlist = parseAllowlist("# a comment\n\n   \n");
    assert.equal(allowlist.files.size, 0);
    assert.equal(allowlist.anywhere.size, 0);
    assert.equal(allowlist.perFile.size, 0);
  });

  it("allows a whole file", () => {
    const allowlist = parseAllowlist("packages/workshop-frontend/src/Api.tsx\n");
    assert.equal(isAllowlisted(allowlist, candidate), true);
  });

  it("allows a whole directory when the entry ends in a slash", () => {
    const allowlist = parseAllowlist("packages/workshop-frontend/src/\n");
    assert.equal(isAllowlisted(allowlist, candidate), true);
    assert.equal(isAllowlisted(parseAllowlist("packages/workshop-backend/\n"), candidate), false);
  });

  it("allows one string in one file, and not the same string elsewhere", () => {
    const allowlist = parseAllowlist("packages/workshop-frontend/src/Api.tsx|GET the thing\n");
    assert.equal(isAllowlisted(allowlist, candidate), true);
    assert.equal(isAllowlisted(allowlist, { ...candidate, file: "other.tsx" }), false);
  });

  it("allows one string anywhere", () => {
    const allowlist = parseAllowlist("*|GET the thing\n");
    assert.equal(isAllowlisted(allowlist, { ...candidate, file: "other.tsx" }), true);
  });

  it("matches on normalized text, so wrapping in the source does not matter", () => {
    const allowlist = parseAllowlist("*|GET   the thing\n");
    assert.equal(isAllowlisted(allowlist, candidate), true);
  });

  it("leaves an unrelated string reported", () => {
    const allowlist = parseAllowlist("*|Something else\n");
    assert.equal(isAllowlisted(allowlist, candidate), false);
  });
});

// A `.ts` file has no JSX, so the detector judges its string and template literals directly. That
// is a far noisier question than "is this JSX text", and these pin the positions ruled out by
// syntax rather than by an allowlist entry.
function tsTexts(source: string): string[] {
  return findStringCandidates("packages/workshop-frontend/src/sample.ts", source)
    .map((candidate) => candidate.text);
}

describe("findStringCandidates: .ts string literals", () => {
  it("finds a sentence and a lone label word", () => {
    assert.deepEqual(tsTexts("throw new Error('The selected file is not an image.')"),
      ["The selected file is not an image."]);
    assert.deepEqual(tsTexts("const status = 'Unchanged'"), ["Unchanged"]);
  });

  it("reports the kind and a span that slices back to the text", () => {
    const source = "const message = 'Choose an image file.'";
    const [candidate] = findStringCandidates("packages/workshop-frontend/src/sample.ts", source);
    assert.equal(candidate.kind, "string-literal");
    assert.equal(source.slice(candidate.textStart, candidate.textEnd), "Choose an image file.");
  });

  it("finds a template literal, standing each placeholder in as {0}", () => {
    assert.deepEqual(tsTexts("const title = `Failed to ${verb} the action`"),
      ["Failed to {0} the action"]);
  });

  it("keeps a templated URL reading as a URL rather than as a sentence", () => {
    // The placeholder carries no space on purpose; joining with one would make this a phrase.
    assert.deepEqual(tsTexts("const u = `https://dash.example.com/?to=/${id}/ai/credits`"), []);
  });

  it("does not report a template's placeholders as separate candidates", () => {
    assert.deepEqual(tsTexts("const t = `Logo request failed with status ${response.status}`"),
      ["Logo request failed with status {0}"]);
  });

  it("skips a type, which is never a value", () => {
    assert.deepEqual(tsTexts("export type Status = 'Added' | 'Deleted' | 'Unchanged'"), []);
  });

  it("skips module specifiers, static and dynamic", () => {
    assert.deepEqual(tsTexts("import { a } from './some other module'"), []);
    assert.deepEqual(tsTexts("export * from './some other module'"), []);
    assert.deepEqual(tsTexts("const m = await import('./some other module')"), []);
  });

  it("skips a property name but keeps the value beside it", () => {
    assert.deepEqual(tsTexts("const o = { 'a phrase key': 'A real label' }"), ["A real label"]);
    assert.deepEqual(tsTexts("const v = o['a phrase key']"), []);
  });

  it("skips console arguments, which are written for whoever reads the log", () => {
    assert.deepEqual(tsTexts("console.error('Failed to load the catalog:', err)"), []);
    assert.deepEqual(tsTexts("console.warn('Retrying the upload')"), []);
  });

  it("skips a needle matched against data", () => {
    for (const method of ["includes", "startsWith", "endsWith", "indexOf", "lastIndexOf"]) {
      assert.deepEqual(tsTexts(`if (message.${method}('Peer closed WebSocket')) fail()`), [],
        method);
    }
  });

  it("skips a CSS class list assigned to className", () => {
    assert.deepEqual(tsTexts("row.className = 'deleted-num-row deleted-omitted-row'"), []);
  });

  it("skips an operand of an equality test, which is matched rather than shown", () => {
    for (const operator of ["===", "!==", "==", "!="]) {
      assert.deepEqual(tsTexts(`if (model.status ${operator} 'Modified') update()`), [], operator);
      assert.deepEqual(tsTexts(`if ('Modified' ${operator} model.status) update()`), [], operator);
    }
  });

  it("keeps a comparison against something that is also rendered", () => {
    // Only the operand is excluded; a label beside it is still copy.
    assert.deepEqual(tsTexts("const label = status === 'Modified' ? 'Changed by you' : ''"),
      ["Changed by you"]);
  });

  it("skips a switch label, which is the same comparison written another way", () => {
    assert.deepEqual(tsTexts("switch (s) { case 'Unchanged': return 0 }"), []);
  });

  it("skips text that is already localized", () => {
    assert.deepEqual(tsTexts("const m = t`Choose an image file.`"), []);
    assert.deepEqual(tsTexts("const m = msg`Choose an image file.`"), []);
    assert.deepEqual(tsTexts("const m = i18n._('Choose an image file.')"), []);
    assert.deepEqual(tsTexts("const m = defineMessage('Choose an image file.')"), []);
  });

  it("skips a plural/select/selectOrdinal message, one level down in its options object", () => {
    // The message text is a property value inside the options object, not a direct call argument,
    // so this needs its own branch rather than the plain call-argument check above.
    assert.deepEqual(
      tsTexts("const m = plural(hidden, { one: 'Show # hidden line', other: 'Show # hidden lines' })"),
      []);
    assert.deepEqual(
      tsTexts("const m = select(kind, { book: 'Book', other: 'Item' })"),
      []);
    assert.deepEqual(
      tsTexts("const m = selectOrdinal(place, { one: '#st', two: '#nd', other: '#th' })"),
      []);
  });

  it("does not let an unrelated call's options object through", () => {
    // Same shape (a string in an object literal passed to a call), but the callee is not one of
    // the localizing macros -- the message must still be reported.
    assert.deepEqual(
      tsTexts("const m = configure({ one: 'Show # hidden line', other: 'Show # hidden lines' })"),
      ["Show # hidden line", "Show # hidden lines"]);
  });

  it("keeps a thrown message, because whether a caller shows it is not a syntactic question", () => {
    assert.deepEqual(tsTexts("throw new Error('Logo source image is too large (max 5 MB).')"),
      ["Logo source image is too large (max 5 MB)."]);
  });

  it("does not apply the string-literal rule to a .tsx file", () => {
    // In a `.tsx` the copy is in JSX text and whitelisted attributes; the literals around it are
    // props, ids and class names, and scanning them would bury the real findings.
    assert.deepEqual(texts("const label = 'A stray sentence here'"), []);
  });
});

describe("isExcludedFile: declaration files", () => {
  it("excludes .d.ts, which holds types and never values", () => {
    assert.equal(isExcludedFile("packages/workshop-frontend/src/vite-env.d.ts"), true);
  });
});

describe("listSourceFiles: extensions", () => {
  it("walks .ts and .tsx by default and can be narrowed", () => {
    const both = listSourceFiles(SCAN_ROOT);
    const onlyTsx = listSourceFiles(SCAN_ROOT, [".tsx"]);
    assert.ok(both.length > onlyTsx.length);
    assert.ok(both.some((path) => path.endsWith(".ts")));
    assert.equal(onlyTsx.every((path) => path.endsWith(".tsx")), true);
    assert.equal(both.some((path) => path.endsWith(".d.ts")), false);
  });
});

// Opt-in, because the codemod has no transform for these yet: reporting them by default would
// change what the guard says without changing what can be fixed mechanically.
function exprCandidates(source: string): StringCandidate[] {
  return findStringCandidates("packages/workshop-frontend/src/Sample.tsx", source,
    { includeJsxExpressions: true });
}

function exprTexts(source: string): string[] {
  return exprCandidates(source).map((candidate) => candidate.text);
}

describe("findStringCandidates: JSX expression literals", () => {
  const TERNARY = "<button>{busy ? 'Enabling...' : 'Always approve'}</button>";

  it("reports them by default, and nothing when asked not to", () => {
    assert.deepEqual(texts(TERNARY), ["Enabling...", "Always approve"]);
    assert.deepEqual(
      findStringCandidates("packages/workshop-frontend/src/Sample.tsx", TERNARY,
        { includeJsxExpressions: false }),
      [],
    );
  });

  it("finds both sides of a ternary rendered as a child", () => {
    assert.deepEqual(exprTexts(TERNARY), ["Enabling...", "Always approve"]);
    assert.equal(exprCandidates(TERNARY).every((c) => c.kind === "jsx-expression"), true);
    assert.equal(exprCandidates(TERNARY)[0].parentTag, "button");
  });

  it("finds a literal behind a logical and", () => {
    assert.deepEqual(exprTexts("<p>{isEmpty && 'No outputs yet'}</p>"), ["No outputs yet"]);
  });

  it("finds a ternary in a whitelisted attribute, and records which one", () => {
    const found = exprCandidates("<i title={busy ? 'Saving changes' : 'Save changes'} />");
    assert.deepEqual(found.map((c) => c.text), ["Saving changes", "Save changes"]);
    assert.equal(found[0].attribute, "title");
  });

  it("ignores an attribute that carries no copy", () => {
    assert.deepEqual(exprTexts("<div className={on ? 'row is-on' : 'row is-off'} />"), []);
    assert.deepEqual(exprTexts("<div onClick={() => track('user opened the menu')} />"), []);
  });

  it("does not report a braced literal twice", () => {
    // `title={"Close"}` is already a jsx-attribute candidate.
    const found = exprCandidates('<i title={"Close preview"} />');
    assert.deepEqual(found.map((c) => c.kind), ["jsx-attribute"]);
  });

  // The attribute collector accepts a quoted string and nothing else, so a template value has to be
  // reported here or it is reported nowhere -- invisible to the codemod and to the guard alike.
  it("finds a template attribute value, which no other collector accepts", () => {
    const found = exprCandidates("<i title={`Actions for ${name}`} />");
    assert.deepEqual(found.map((c) => c.text), ["Actions for {0}"]);
    assert.equal(found[0].kind, "jsx-expression");
    assert.equal(found[0].attribute, "title");
  });

  it("finds a template attribute value with no substitution in it", () => {
    assert.deepEqual(exprTexts("<i title={`Close preview`} />"), ["Close preview"]);
  });

  it("finds a template attribute value on a component prop", () => {
    const found = exprCandidates("<Tooltip content={`Used the gadget at ${when}`} />");
    assert.deepEqual(found.map((c) => c.text), ["Used the gadget at {0}"]);
    assert.equal(found[0].attribute, "content");
  });

  it("keeps a template that builds a URL or a path out", () => {
    assert.deepEqual(exprTexts("<i title={`https://dash.cloudflare.com/${account}/ai`} />"), []);
    assert.deepEqual(exprTexts("<i title={`/workspaces/${id}`} />"), []);
  });

  it("keeps a template on an attribute that carries no copy out", () => {
    assert.deepEqual(exprTexts("<div className={`row ${on ? 'is-on' : 'is-off'}`} />"), []);
    assert.deepEqual(exprTexts("<a href={`https://example.com/${id}`} />"), []);
  });

  it("skips what is already inside a Lingui element", () => {
    assert.deepEqual(exprTexts("<Trans>{busy ? 'Enabling...' : 'Always approve'}</Trans>"), []);
  });

  it("skips rendered children of a technical-literal element", () => {
    assert.deepEqual(exprTexts("<code>{on ? 'Enabling...' : 'Always approve'}</code>"), []);
  });

  it("keeps identifiers, enums and paths out", () => {
    assert.deepEqual(exprTexts("<p>{on ? nameA : nameB}</p>"), []);
    assert.deepEqual(exprTexts("<p>{on ? 'asc' : 'desc'}</p>"), []);
    assert.deepEqual(exprTexts("<p>{on ? 'onClick' : 'onHover'}</p>"), []);
    assert.deepEqual(exprTexts("<p>{on ? '/workspaces/new' : '/workspaces'}</p>"), []);
  });

  it("still finds ordinary JSX text alongside", () => {
    assert.deepEqual(exprTexts("<p>Delete this workspace? {busy && 'Working on it'}</p>"),
      ["Delete this workspace?", "Working on it"]);
  });

  it("keeps an operand of an equality test out", () => {
    assert.deepEqual(exprTexts("<div>{model.status !== 'Modified' && <b>{label}</b>}</div>"), []);
  });
});
