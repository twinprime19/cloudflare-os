# Residue Extraction Conventions

You are wrapping user-facing English strings in a React + TypeScript frontend with Lingui v6 macros.
A codemod already wrapped 654 of 745 strings mechanically. You are handling part of the 91 it
refused, because they need judgement.

Everything below was verified by running `lingui extract` against this exact toolchain. The
"after" code in every example is real code from this repository. Copy the shapes.

---

## 0. Hard rules — read these first

1. **Worktree:** `~/repos/cloudflare-os-i18n-core`. All paths in this document and in the
   residue report are relative to that directory. Do not touch the main checkout at
   `~/repos/cloudflare-os`.
2. **Only edit the files in your assigned shard.** Your shard is the explicit file list in the
   prompt that launched you. Another agent owns every other file, and several agents are working at
   the same time. If a fix seems to require editing a file outside your shard, stop and escalate it.
3. **Never start a dev server.** Never run `pnpm dev`, `vite`, or `run-local`.
4. **Never commit, stage, push, or create branches.**
5. **Never hand-edit `packages/i18n/locales/en.po`.** It is generated. Do not run
   `pnpm run i18n:extract` either — it rewrites a shared file and the orchestrator runs it once per
   wave.
6. **Never read or grep inside `node_modules`.** A repo hook blocks it.
7. **Do not reformat, re-indent, rename, or "improve" anything you are not wrapping.** No prettier,
   no lint --fix, no comment tidying. Your diff should contain nothing but wraps.
8. **Do not change the rendered English by even one character.** Not the wording, not the
   punctuation, not a space. You are not a copy editor.
9. **Scope = the residue entries assigned to you.** A bare string you happen to notice nearby that
   is not in the report is not yours (see §7.2).

---

## 1. THE CARDINAL RULE — never produce sentence fragments

A sentence that a user reads as one sentence must be **one message**, with placeholders for the
parts that vary. Never split it into two or three messages.

Translators receive a list of messages with no surrounding code. A fragment like `"Shared by"` is
untranslatable: in many languages the word order, the case ending, and the preposition all depend on
what follows. A translator cannot reorder pieces they receive separately. This is the single most
damaging mistake available to you, and it is invisible in English — the app will look perfect and
the localization will be broken.

### WRONG — three fragments

```tsx
<p>
  <Trans>Shared by</Trans> {gadget.owner.name}
</p>
```

Produces messages `Shared by` — a dangling fragment.

### RIGHT — one message with a placeholder

```tsx
<p>
  <Trans>Shared by {gadget.owner.name}</Trans>
</p>
```

Produces message `Shared by {0}`, with `#. placeholder {0}: gadget.owner.name` recorded for the
translator.

The same rule in prose: **if you can read the pieces aloud as one sentence, they are one message.**

---

## 2. How to read `codemod-residue-report.json`

Path: `~/repos/cloudflare-os/plans/260818-1401-i18n-frontend-string-extraction/codemod-residue-report.json`

```jsonc
{
  "files": [
    {
      "file": "packages/workshop-frontend/src/ShareModal.tsx",   // repo-relative, forward slashes
      "residue": [
        {
          "kind": "jsx-text",         // or "jsx-attribute"
          "parentTag": "p",           // jsx-text only; "" means the parent is a <> fragment
          "attribute": "placeholder", // jsx-attribute only
          "text": "People with",      // whitespace-collapsed; may be a FRAGMENT of the sentence
          "line": 256,                // 1-based
          "column": 15,               // 1-based
          "reason": "split-by-markup",
          "group": "packages/workshop-frontend/src/ShareModal.tsx:8944"
        }
      ]
    }
  ]
}
```

Four things that matter:

- **`line`/`column` are post-apply.** They are correct against the current file contents in the
  worktree. Open the file at that line and you will find the text.
- **`text` is often a fragment, not the message.** It is what the detector saw between two
  expressions. The message you write is usually longer than `text`.
- **`group` is the unit of work.** Every entry sharing a `group` value is a piece of **one** message
  and must be fixed in **one** edit. 91 entries collapse into ~70 tasks. If your shard has three
  entries with the same `group`, that is one `<Trans>`, not three.
- **`reason` is structural, not semantic:**

  | reason | what the codemod hit | what you usually do |
  |---|---|---|
  | `interpolated` | a sibling `{expr}` interrupts the text | one `<Trans>` over the whole element, expression inline |
  | `split-by-markup` | a sibling element renders text of its own | one `<Trans>` over the whole element; inline tags become `<0>`, `<1>` |
  | `attribute-unsafe-value` | the attribute value has escape sequences | §6.3 |

**Work bottom-up inside each file.** Fix the highest line number first, then the next highest. Your
edits shift the lines below them; going upward keeps every remaining line number in the report valid.

---

## 3. The four wrapping forms

### 3.1 `<Trans>` — for text rendered as JSX children

```tsx
import { Trans } from '@lingui/react/macro'

<p><Trans>Delete workspace?</Trans></p>
```

Anything inside it becomes part of one message: plain text, `{expressions}` (placeholders), and
nested elements (`<0>…</0>` tags).

### 3.2 `` t`…` `` — for attribute values and plain strings

```tsx
import { t } from '@lingui/core/macro'

<input placeholder={t`Search blueprints…`} />
```

This is the **core** macro on the global i18n instance. There is no hook to call and no provider to
add — the app already mounts one at `main.tsx`. Do not import `t` from `@lingui/macro`, do not use
`useLingui()`, do not use `i18n._()`.

### 3.3 `<Plural>` — for counts

```tsx
import { Plural } from '@lingui/react/macro'

<Plural value={count} one="# resource hidden" other="# resources hidden" />
```

`#` renders `value`. `<Plural>` **can be nested inside `<Trans>`** (verified) when the count sits
inside a longer sentence — see §5, E7. English needs exactly the two categories `one` and `other`;
do not add `zero`, `few`, or `many`.

### 3.4 Interpolation — the answer to concatenation

Put the expression inside the message. Simple identifiers keep their name (`{actionLabel}`);
property accesses and calls become numbered (`{0}`, `{1}`) with a `#. placeholder` comment. Both are
fine. Do **not** extract an expression into a new `const` to get a prettier placeholder name — that
is a code change you are not authorized to make.

### 3.5 Imports

Match the file's own quote and semicolon style — look at the last import in the file. Most files use
single quotes and no semicolons; `ChatInterface.tsx` and `SignupPage.tsx` use double quotes and
semicolons. Put macro imports after the last package import, before the relative ones.

```tsx
import { t } from '@lingui/core/macro'
import { Plural, Trans } from '@lingui/react/macro'
```

**Every residue file except one already imports what it needs.** These are the only imports you have
to add:

| file | add |
|---|---|
| `src/AdminPage.tsx` | new line `import { t } from '@lingui/core/macro'`; change the `@lingui/react/macro` import to `{ Plural, Trans }` |
| `src/ShareModal.tsx` | change to `import { Plural, Trans } from '@lingui/react/macro'` |
| `src/components/ConnectConnectorModal.tsx` | change to `import { Plural, Trans } from '@lingui/react/macro'` |
| `src/components/billing/UsageSettings.tsx` | change to `import { Plural, Trans } from '@lingui/react/macro'` |
| `src/components/billing/OutOfCreditsModal.tsx` | change to `import { Plural, Trans } from '@lingui/react/macro'` |
| `src/components/auth/OAuthButtons.tsx` | has no Lingui import at all; add `import { Trans } from '@lingui/react/macro'` after line 4 (`@cloudflare/kumo`) |

`noUnusedLocals` is on. Only add an import you actually use, or the typecheck fails.

---

## 4. Whitespace — the second most common way to break this

JSX renders whitespace, so a lost or added space is a visible bug.

- A newline plus indentation between two JSX lines renders as **one space**. Keep the line breaks
  where they were.
- `{' '}` renders one space. Inside `<Trans>` it is absorbed into the message as a plain space
  (verified). Keep every `{' '}` exactly where it was.
- A space between `</Trans>` and the next expression is significant. Do not delete it.
- HTML entities (`&apos;`, `&ldquo;`, `&rdquo;`, `&rsquo;`) stay **verbatim in the source**. Lingui
  decodes them during extraction, so `You can&apos;t undo this.` becomes `You can't undo this.` in
  the catalog. Do not convert entities to characters yourself.

---

## 5. Worked examples

Every one is a real residue entry. Line numbers are the ones in the report.

### E1 — Concatenation, simplest form
`src/components/RecentApps.tsx:57` (identical at `src/components/GadgetList.tsx:116`)

```tsx
// BEFORE
<p className="text-xs text-kumo-subtle truncate mt-0.5">
  Shared by {gadget.owner.name}
</p>

// AFTER
<p className="text-xs text-kumo-subtle truncate mt-0.5">
  <Trans>Shared by {gadget.owner.name}</Trans>
</p>
```

Message: `Shared by {0}`

Same shape, same fix: `src/ChatInterface.tsx:6348` (`Connect {msg.vendorName}`),
`src/ChatInterface.tsx:6828` (`Agent · {chat.spawnerName}`), `src/ChatInterface.tsx:7186`
(`Switched to {entry.author.name}`), `src/ChatInterface.tsx:7140`
(`{entry.requestedBy.name} compacted the context`), `src/BlueprintLandingPage.tsx:817`
(`Updated {…}`), `src/BlueprintLandingPage.tsx:1017` (`Configure {…}`),
`src/BlueprintLandingPage.tsx:1143` (`Screenshot of {title}`), `src/BlueprintLandingPage.tsx:1640`
(`The required resource type for this binding isn't offered by {…}.`), `src/BlueprintModal.tsx:597`
(`Delete "{bp.title}"?`), `src/Connections.tsx:257`, `src/Connections.tsx:389`,
`src/OnboardingWizard.tsx:788`, `src/ShareModal.tsx:760`, `src/ShareModal.tsx:852`,
`src/ShareModal.tsx:1056`, `src/routes/outputs.tsx:206`, `src/routes/outputs.tsx:355`.

**E1 beats §6.3 when the separator sits inside the same text node.** `src/ShareModal.tsx:1056` is
`Created by {sk.createdBy.name} · {formatRelativeTime(sk.created)}` — a label, an interpolation, a
`·` separator, and a helper's output, all in one run. Wrap the whole run:

```tsx
// AFTER
<p className="truncate text-[12px] leading-[15px] tracking-[-0.15px] text-kumo-subtle"><Trans>Created by {sk.createdBy.name} · {formatRelativeTime(sk.created)}</Trans></p>
```

Message: `Created by {0} · {1}`. §6.3 keeps a separator plus data *outside* the message, but §6.3's
cases are appended by a **separate expression** (` — {usage.accountName}`, `` {cond ? ` -- ${x}` : ''} ``)
that the sentence does not need. Here the separator is literal text between two interpolations in
one run, so hoisting it out would manufacture two fragments — the worse failure. One message wins:
it also lets a translator reorder the phrase and localize the separator.

### E2 — Concatenation split by markup (three residue entries, one message)
`src/components/AutoApproveConfirmDialog.tsx:47–49` — group `…:1622`

```tsx
// BEFORE
<Dialog.Description className="mt-1 text-[12px] leading-4 font-normal tracking-[-0.2px] text-kumo-subtle">
  Future <span className="font-medium text-kumo-default">{actionLabel}</span> actions on{' '}
  <span className="font-medium text-kumo-default">{resourceTitle}</span> will be applied
  automatically, without asking for approval. This action will be applied now too.
</Dialog.Description>

// AFTER
<Dialog.Description className="mt-1 text-[12px] leading-4 font-normal tracking-[-0.2px] text-kumo-subtle">
  <Trans>
    Future <span className="font-medium text-kumo-default">{actionLabel}</span> actions on{' '}
    <span className="font-medium text-kumo-default">{resourceTitle}</span> will be applied
    automatically, without asking for approval. This action will be applied now too.
  </Trans>
</Dialog.Description>
```

Message: `Future <0>{actionLabel}</0> actions on <1>{resourceTitle}</1> will be applied automatically, without asking for approval. This action will be applied now too.`

The wrap goes **around the whole run**, outermost. The `<span>`s, their `className`s, and the
`{' '}` are untouched.

Same shape: `src/Connections.tsx:260`, `src/components/chat/PermissionToast.tsx:35`,
`src/components/chat/PermissionToast.tsx:155` (message starts with markup — that is fine),
`src/BlueprintLandingPage.tsx:1664`, `src/components/BlueprintBindingCard.tsx:75`,
`src/Connections.tsx:323`, `src/components/billing/UsageSettings.tsx:107`
(`<Trans>Resets at 00:00 UTC, in{' '}<ResetCountdown … />.</Trans>` — the countdown component
becomes `<0/>`).

### E3 — The inner `<Trans>` the codemod already added must be absorbed
`src/LoginPage.tsx:154–159`

```tsx
// BEFORE  (the codemod wrapped the link text and left the sentence around it bare)
<p className="text-center text-sm text-kumo-subtle mt-6">
  Don't have an account?{' '}
  <Link to="/signup" className="text-kumo-brand hover:underline font-medium">
    <Trans>Create one</Trans>
  </Link>
</p>

// AFTER  (one message; the inner <Trans> is removed, not nested)
<p className="text-center text-sm text-kumo-subtle mt-6">
  <Trans>
    Don't have an account?{' '}
    <Link to="/signup" className="text-kumo-brand hover:underline font-medium">
      Create one
    </Link>
  </Trans>
</p>
```

Message: `Don't have an account? <0>Create one</0>`

Whenever the residue text sits next to an element that already contains a `<Trans>`, the fix is to
hoist one `<Trans>` to the top and delete the inner one.

Same shape: `src/SignupPage.tsx:219`, `src/components/billing/OutOfCreditsModal.tsx:161`,
`src/components/billing/UsageSettings.tsx:202`, `src/routes/providers.tsx:256`.

### E4 — Inline technical element inside a sentence
`src/GadgetEditor.tsx:1679`

```tsx
// BEFORE
<div className="rounded-full border border-kumo-line bg-kumo-base/90 px-4 py-1.5 text-[13px] leading-[18px] text-kumo-default shadow-md backdrop-blur-sm">
  Press <kbd className="rounded border border-kumo-line bg-kumo-elevated px-1.5 py-0.5 text-[11px] font-medium">Esc</kbd> to exit full screen
</div>

// AFTER
<div className="rounded-full border border-kumo-line bg-kumo-base/90 px-4 py-1.5 text-[13px] leading-[18px] text-kumo-default shadow-md backdrop-blur-sm">
  <Trans>Press <kbd className="rounded border border-kumo-line bg-kumo-elevated px-1.5 py-0.5 text-[11px] font-medium">Esc</kbd> to exit full screen</Trans>
</div>
```

Message: `Press <0>Esc</0> to exit full screen`

`Esc` on its own is a technical literal you would never wrap (§7). But when it is a word in a
sentence you are wrapping, it rides along as `<0>` content. That is correct and unavoidable.

### E5 — A `<>` fragment passed as a prop becomes `<Trans>`
`src/FileSidebar.tsx:261`

```tsx
// BEFORE
description={<>This removes <span className="font-mono text-kumo-default">{deletingFile}</span> from the gadget. You can&apos;t undo this.</>}

// AFTER
description={<Trans>This removes <span className="font-mono text-kumo-default">{deletingFile}</span> from the gadget. You can&apos;t undo this.</Trans>}
```

Message: `This removes <0>{deletingFile}</0> from the gadget. You can't undo this.`

The fragment `<>…</>` is **replaced** by `<Trans>…</Trans>`, not wrapped inside it. `<Trans>` is
already a single element, so the prop still receives one node.

Same shape: `src/GadgetEditor.tsx:1787`, `src/ChatInterface.tsx:7977`.

### E6 — Manual pluralization → ICU, standalone
`src/AdminPage.tsx:947`

```tsx
// BEFORE
<p className="text-xs text-kumo-subtle px-3 py-1">
  {vendor.resources.length} resource{vendor.resources.length === 1 ? '' : 's'} hidden while disabled.
</p>

// AFTER
<p className="text-xs text-kumo-subtle px-3 py-1">
  <Plural
    value={vendor.resources.length}
    one="# resource hidden while disabled."
    other="# resources hidden while disabled."
  />
</p>
```

Message: `{0, plural, one {# resource hidden while disabled.} other {# resources hidden while disabled.}}`

**Check the English renders identically before you move on.** `n = 1` → `1 resource hidden while
disabled.` `n = 4` → `4 resources hidden while disabled.` Both match the old code exactly. If your
`one`/`other` strings do not reproduce the old output character for character, you got it wrong.

Same shape: `src/components/ConnectConnectorModal.tsx:349`
(`one="# resource to add"` / `other="# resources to add"`).

### E7 — Plural inside a longer sentence
`src/components/billing/UsageSettings.tsx:102–104`

```tsx
// BEFORE
<p className="text-sm text-kumo-default">
  {usage.remaining} of {usage.dailyLimit}{' '}
  {usage.dailyLimit === 1 ? 'request' : 'requests'} remaining today
</p>

// AFTER
<p className="text-sm text-kumo-default">
  <Trans>
    {usage.remaining} of{' '}
    <Plural value={usage.dailyLimit} one="# request" other="# requests" /> remaining today
  </Trans>
</p>
```

Message: `{0} of {1, plural, one {# request} other {# requests}} remaining today`

Note what happened to `{usage.dailyLimit}`: it disappeared from the text because `#` inside the
`<Plural>` renders it. Renders `3 of 10 requests remaining today`, same as before. One message, not
three.

### E8 — Plural with another value interpolated
`src/ShareModal.tsx:1001`

```tsx
// BEFORE
<p className="mb-1.5 text-[12px] leading-4 text-kumo-subtle">
  {downstreamDependents.length} other {downstreamDependents.length === 1 ? 'person loses' : 'people lose'} access through {profile.name}. Keep anyone?
</p>

// AFTER
<p className="mb-1.5 text-[12px] leading-4 text-kumo-subtle">
  <Trans>
    <Plural
      value={downstreamDependents.length}
      one="# other person loses"
      other="# other people lose"
    />{' '}
    access through {profile.name}. Keep anyone?
  </Trans>
</p>
```

Message: `{0, plural, one {# other person loses} other {# other people lose}} access through {1}. Keep anyone?`

**Put only the words that vary by count in `one`/`other`; nest the `<Plural>` in a `<Trans>` that owns
the rest of the sentence.** Do *not* reach for template-literal props
(``one={`# other person loses access through ${profile.name}. Keep anyone?`}``). That shape renders
correctly, but Lingui numbers every template-literal expression independently, so the same variable
lands in the catalog twice — verified:

```
#. placeholder {0}: downstreamDependents.length
#. placeholder {1}: profile.name
#. placeholder {2}: profile.name
msgid "{0, plural, one {# other person loses access through {1}. …} other {# other people lose access through {2}. …}}"
```

A translator then has to work out that `{1}` and `{2}` are the same person's name, and that the
number to type differs by branch. (Both values are passed at runtime, so it still renders — the cost
is entirely the translator's.) The nested form emits one `#. placeholder {1}: profile.name` and one
`{1}`. Both forms render byte-identically in both plural categories; prefer the one the translator
can read.

Related: `src/ShareModal.tsx:1108` has **no** interpolated value, so there is nothing to duplicate
and the standalone E6 shape is right —
`one="# person loses access through this link. Keep anyone?"` /
`other="# people lose access through this link. Keep anyone?"`. The rule is about interpolated
values, not about plurals in general: a plural whose branches differ only in wording may keep whole
sentences in `one`/`other`.

### E9 — Conditional where both branches share a tail
`src/ShareModal.tsx:255–257`

```tsx
// BEFORE  (the tail " prove their own account can reach:" is outside the ternary)
<p className="text-[12px] leading-[16px] tracking-[-0.15px] text-kumo-subtle">
  {role ? (
    <>People with <span className="font-medium text-kumo-default">{roleLabel(role)}</span> access must</>
  ) : 'Recipients must'} prove their own account can reach:
</p>

// AFTER  (the tail is duplicated into both branches so each branch is a whole sentence)
<p className="text-[12px] leading-[16px] tracking-[-0.15px] text-kumo-subtle">
  {role ? (
    <Trans>
      People with <span className="font-medium text-kumo-default">{roleLabel(role)}</span> access
      must prove their own account can reach:
    </Trans>
  ) : (
    <Trans>Recipients must prove their own account can reach:</Trans>
  )}
</p>
```

Messages: `People with <0>{0}</0> access must prove their own account can reach:` and
`Recipients must prove their own account can reach:`

**Duplicating the shared tail is correct.** Two complete sentences beat one sentence assembled from
parts. Never leave the tail outside the branches.

### E10 — Conditional that picks a word inside a sentence
`src/ChatInterface.tsx:6749`

```tsx
// BEFORE
<p className="text-[13px] leading-[18px] text-kumo-inactive">
  No conversations started by {chatListScope === "agents" ? "agents" : "people"} yet
</p>

// AFTER
<p className="text-[13px] leading-[18px] text-kumo-inactive">
  {chatListScope === "agents" ? (
    <Trans>No conversations started by agents yet</Trans>
  ) : (
    <Trans>No conversations started by people yet</Trans>
  )}
</p>
```

Two whole sentences. **Never** `<Trans>No conversations started by {cond ? "agents" : "people"} yet</Trans>` —
that leaves the nouns untranslated and gives the translator an opaque `{0}`.

### E11 — Conditional that appends a clause
`src/BlueprintModal.tsx:349–350`

```tsx
// BEFORE
<p className="m-0 mt-0.5 text-[12px] leading-4 font-normal tracking-[-0.2px] text-kumo-subtle">
  Optional image shown on Explore and the blueprint detail page.
  {formMode === 'edit' && !newScreenshotUrl && editingBlueprint?.screenshotUrl && !clearScreenshot ? ' The current screenshot will stay unless you upload a new one.' : ''}
</p>

// AFTER
<p className="m-0 mt-0.5 text-[12px] leading-4 font-normal tracking-[-0.2px] text-kumo-subtle">
  <Trans>Optional image shown on Explore and the blueprint detail page.</Trans>
  {formMode === 'edit' && !newScreenshotUrl && editingBlueprint?.screenshotUrl && !clearScreenshot ? (
    <>
      {' '}
      <Trans>The current screenshot will stay unless you upload a new one.</Trans>
    </>
  ) : (
    ''
  )}
</p>
```

Messages: `Optional image shown on Explore and the blueprint detail page.` and
`The current screenshot will stay unless you upload a new one.`

Two separate **sentences**, so two messages is right — this is not fragmenting. The leading space
that used to live inside the string literal (`' The current…'`) moves out to `{' '}` so the message
itself has no stray leading space.

Same shape, `src/BlueprintLandingPage.tsx:1086` — a whole sentence, then a ternary between two whole
sentences:

```tsx
// AFTER
<Dialog.Description className="mt-2 text-kumo-subtle">
  <Trans>Delete "{blueprint?.metadata.title}"?</Trans>{' '}
  {canDeleteOwnedBlueprint ? (
    <Trans>This blueprint link will stop working, but gadgets already created from it won’t be affected.</Trans>
  ) : (
    <Trans>This blueprint was uploaded manually and cannot be recovered.</Trans>
  )}
</Dialog.Description>
```

The single space that used to sit between `"?` and `{canDeleteOwnedBlueprint` becomes `{' '}`.

### E12 — Conditional clause in the middle of a sentence
`src/routes/outputs.tsx:701–707`

```tsx
// BEFORE
description={
  <>
    This permanently removes the output from “{removeOutput?.workspaceTitle}”
    {removeOutput?.owner ? ', for everyone with access to that workspace' : ''}. Other
    outputs in that workspace stay available. This can’t be undone.
  </>
}

// AFTER
description={
  removeOutput?.owner ? (
    <Trans>
      This permanently removes the output from “{removeOutput?.workspaceTitle}”, for everyone with
      access to that workspace. Other outputs in that workspace stay available. This can’t be
      undone.
    </Trans>
  ) : (
    <Trans>
      This permanently removes the output from “{removeOutput?.workspaceTitle}”. Other outputs in
      that workspace stay available. This can’t be undone.
    </Trans>
  )
}
```

Messages: `This permanently removes the output from “{0}”, for everyone with access to that
workspace. Other outputs in that workspace stay available. This can’t be undone.` and the same
without the clause.

The ternary moves **outward** so each branch is a complete message. Copy the curly quotes `“ ” ’`
exactly — they are not ASCII quotes.

Same shape, `src/components/billing/UsageSettings.tsx:179` — a label whose value is a ternary, one
branch of which is itself copy (`unknown`):

```tsx
// AFTER
<p className="text-sm text-kumo-default">
  {usage.balance !== null ? (
    <Trans>Account balance: <strong>${usage.balance.toFixed(2)}</strong></Trans>
  ) : (
    <Trans>Account balance: <span className="text-kumo-subtle">unknown</span></Trans>
  )}
</p>
```

A label and its value belong in one message: `Account balance:` alone is a fragment, and `unknown`
alone is a fragment.

### E13 — Mixed JSX text and an inline expression that is part of the sentence
`src/components/chat/DataTab.tsx:103`

```tsx
// BEFORE
<span className="font-mono text-xs text-kumo-subtle">
  {sampleDataRows.reduce((sum, r) => sum + r.messages, 0).toLocaleString()} total messages
</span>

// AFTER
<span className="font-mono text-xs text-kumo-subtle">
  <Trans>
    {sampleDataRows.reduce((sum, r) => sum + r.messages, 0).toLocaleString()} total messages
  </Trans>
</span>
```

Message: `{0} total messages`. A complicated expression is fine inside `<Trans>` — it becomes `{0}`
and Lingui records what it was. Leave the expression exactly as written.

### E13b — NEAR MISS: a count that must NOT become a plural
`src/components/chat/DataTab.tsx:100`, ten lines above E13

```tsx
// BEFORE
<span className="font-mono text-xs text-kumo-subtle">
  {sampleDataRows.length} rows in channels
</span>

// AFTER  — plain interpolation, NOT <Plural>
<span className="font-mono text-xs text-kumo-subtle">
  <Trans>{sampleDataRows.length} rows in channels</Trans>
</span>
```

There is a number and a plural noun, but the current code has **no** `? 's' : ''` and **no**
`? 'row' : 'rows'`. It renders `1 rows in channels` today. Converting it to `<Plural>` would render
`1 row in channels` — a copy change you are not allowed to make. **Only convert to `<Plural>` when
the source already contains manual singular/plural logic.**

### E14 — Self-closing element inside the sentence
`src/ChatInterface.tsx:3874`

```tsx
// BEFORE
<p className="mt-2 border-t border-kumo-line pt-2 text-[11px] leading-[15px] tracking-[-0.1px] text-kumo-inactive">
  Use the <ArrowUUpLeft size={12} className="mx-0.5 inline-block align-[-2px]" aria-hidden="true" /><span className="sr-only"><Trans>undo arrow</Trans></span> under any agent response to discard from that turn onward.
</p>

// AFTER
<p className="mt-2 border-t border-kumo-line pt-2 text-[11px] leading-[15px] tracking-[-0.1px] text-kumo-inactive">
  <Trans>Use the <ArrowUUpLeft size={12} className="mx-0.5 inline-block align-[-2px]" aria-hidden="true" /><span className="sr-only">undo arrow</span> under any agent response to discard from that turn onward.</Trans>
</p>
```

Message: `Use the <0/><1>undo arrow</1> under any agent response to discard from that turn onward.`

The icon becomes a self-closing `<0/>` placeholder. The inner `<Trans>` is absorbed (E3). There must
be **no** space added between `/>` and `<span` — there is none today.

### E15 — `attribute-unsafe-value`: escape sequences in a placeholder
`src/AdminPage.tsx:618` and `src/AdminPage.tsx:698`

```tsx
// BEFORE (line 618)
placeholder={'e.g. 🎉 New: blueprints now support imports — [learn more](https://example.com).'}

// AFTER
placeholder={t`e.g. 🎉 New: blueprints now support imports — [learn more](https://example.com).`}
```

```tsx
// BEFORE (line 698)
placeholder={'e.g. Heads up: scheduled maintenance Saturday — see [status](https://status.example.com).'}

// AFTER
placeholder={t`e.g. Heads up: scheduled maintenance Saturday — see [status](https://status.example.com).`}
```

The `\uXXXX` escapes are replaced with the literal characters (`🎉` = 🎉, `—` = —).
Verified: both extract to exactly the string the old code produced. `AdminPage.tsx` needs a new
`import { t } from '@lingui/core/macro'` (§3.5).

The third one, `src/AdminPage.tsx:753`, contains a `\n` and is a **mandatory escalation** (§8.1).

### E16 — A backend-originated value inside a wrapped label
`src/routes/__root.tsx:68`

```tsx
// BEFORE
<p className="text-sm text-kumo-danger">Authentication error: {error}</p>

// AFTER
<p className="text-sm text-kumo-danger"><Trans>Authentication error: {error}</Trans></p>
```

Message: `Authentication error: {error}`. The label is copy and gets wrapped. The value in `{error}`
comes from the server and is rendered verbatim — it becomes a placeholder and is **never** wrapped.

### E17 — A brand name inside a wrapped sentence
`src/components/auth/OAuthButtons.tsx:116`

```tsx
// BEFORE
Continue with {vendor.displayName}

// AFTER
<Trans>Continue with {vendor.displayName}</Trans>
```

Message: `Continue with {0}`. `vendor.displayName` holds `Google`, `Microsoft`, `GitHub` etc. at
runtime — it stays a placeholder. This file needs the `Trans` import added (§3.5).

Same reasoning: `src/components/ConnectConnectorModal.tsx:322`
(`Gatekeeper sits between {vendorDescription.displayName} and your Gadgets.`) and
`src/components/ConnectConnectorModal.tsx:345`
(`Disconnect {vendorDescription.displayName}? Gadgets using this will lose access.`) — both are one
`<Trans>` each.

---

## 6. Special cases you will meet

### 6.1 The residue `text` is only part of the message
`src/AdminPage.tsx:447` reports `"Shown next to the logo in the top bar. Leave empty to use the
default (&ldquo;"` and `:448` reports `"&rdquo;). Applies on each user&rsquo;s next connection."`.
Same group. One wrap around the whole paragraph, entities untouched:

```tsx
<p className="text-sm text-kumo-subtle mb-5">
  <Trans>Shown next to the logo in the top bar. Leave empty to use the default
  (&ldquo;{DEFAULT_SITE_NAME}&rdquo;). Applies on each user&rsquo;s next connection.</Trans>
</p>
```

### 6.2 Two complete sentences in one text run
`src/BlueprintLandingPage.tsx:1632` is
`The "{binding.gatekeeperName}" gatekeeper is not available on this workshop, so this connection can't be configured.`
— one text run, one `<Trans>`, even though it reads as two clauses. Do not split what the source did
not split.

**Precedence with E11: two *clauses* may not be split; two complete *sentences* may.** What the ban
above protects is the inside of a sentence — that is §1, and it never bends. But when a single
text node contains two whole sentences and a conditional or plural attaches to each one
independently, split at the sentence boundary and let `{' '}` carry the space that used to sit
between them. `src/components/billing/OutOfCreditsModal.tsx:121–134` is the case:

```tsx
// BEFORE — one text node spans "…to continue. Add credits…", with a conditional on each side
<p className="text-sm text-kumo-subtle">
  Your Cloudflare account is connected
  {usage.balance !== null && (<> with a balance of <strong>${usage.balance.toFixed(2)}</strong></>)}
  , but it's below the minimum needed to continue. Add credits to your AI Gateway to
  keep building now
  {usage.resetAt ? (<>{' '}or wait — your free {usage.dailyLimit === 1 ? 'request resets' : 'requests reset'} at
    00:00 UTC, in <ResetCountdown resetAt={usage.resetAt} onElapsed={refresh} />.</>) : '.'}
</p>

// AFTER — sentence 1 owns the balance conditional, sentence 2 owns the reset conditional
<p className="text-sm text-kumo-subtle">
  {usage.balance !== null ? (
    <Trans>
      Your Cloudflare account is connected with a balance of{' '}
      <strong>${usage.balance.toFixed(2)}</strong>, but it's below the minimum needed to
      continue.
    </Trans>
  ) : (
    <Trans>
      Your Cloudflare account is connected, but it's below the minimum needed to
      continue.
    </Trans>
  )}{' '}
  {usage.resetAt ? (
    <Trans>
      Add credits to your AI Gateway to keep building now or wait — your free{' '}
      <Plural value={usage.dailyLimit} one="request resets" other="requests reset" /> at
      00:00 UTC, in <ResetCountdown resetAt={usage.resetAt} onElapsed={refresh} />.
    </Trans>
  ) : (
    <Trans>Add credits to your AI Gateway to keep building now.</Trans>
  )}
</p>
```

Four short whole-sentence messages, versus the 2×2 cross-product of four long near-duplicates you
get by keeping one message — which translators would have to keep in sync by hand. The `&&` becomes
a `?:` so the "no balance" case is a whole sentence rather than an implicit fall-through; `with a
balance of $X` is never emitted on its own. The test is mechanical: **can each piece be read aloud
as a complete sentence?** If yes, splitting is allowed. If any piece is a clause, a label, or a verb
phrase, it is a fragment and §1 forbids it.

### 6.3 A sentence followed by an unrelated data expression
`src/components/format/AdminFormatsPanel.tsx:384`

```tsx
// AFTER
<span className="block">
  <Trans>“{format.output.noun}” — a standard format on this deployment</Trans>
  {format.agentHint ? ` -- ${format.agentHint}` : ''}
</span>
```

The trailing ternary appends agent-supplied data, not copy. Wrap the sentence, leave the ternary
alone.

Same shape, `src/components/billing/UsageSettings.tsx:174` — a lone label followed by an appended
account name:

```tsx
// AFTER
<span>
  <Trans>Connected</Trans>
  {usage.accountName && <> — {usage.accountName}</>}
</span>
```

` — {usage.accountName}` is a separator plus data, not a sentence. It stays outside.

**Scope of this rule:** it applies when the separator plus data live in their own **expression**
alongside the sentence. When the separator is literal text *inside* the same text node, between two
interpolations, E1 wins and the whole run is one message — see the `src/ShareModal.tsx:1056` ruling
at the end of E1.

### 6.4 A standalone sentence next to an English-only expression you must not touch
`src/ChatInterface.tsx:7112` and `src/routes/providers.tsx:271`. Wrap only the residue sentence:

```tsx
// src/ChatInterface.tsx — AFTER
<p className="mb-3 text-[12px] leading-[17px] text-kumo-subtle">
  <Trans>The agent reads this in place of everything earlier in the chat.</Trans>{" "}
  {kept === 0
    ? "Nothing after it was kept."
    : `The ${kept === 1 ? "message" : `${kept} messages`} after the cut ${kept === 1 ? "was" : "were"} kept in full.`}
</p>
```

The ternary below builds an English sentence out of template literals. It is **not** in the residue
report and is **not** yours. Wrap the entry, leave the ternary byte-identical, and add an
`adjacent:` note to your escalation list (§8.3).

---

## 7. What NOT to touch

### 7.1 Never wrap these

| category | example in this repo |
|---|---|
| identifiers, keys, enum values | `gk.name`, `binding.gatekeeperName`, `chatListScope === "agents"` as a *condition* |
| URLs and paths | `https://developers.cloudflare.com/ai-gateway/...`, `to="/signup"`, `src={vendor.logo.url}` |
| CSS class names | every `className="…"`, `DANGEROUS_className` |
| error `code` values | `code: 'not_found'` and anything compared with `===` |
| backend strings rendered verbatim | `{error}`, `{a.accountName}`, `{format.blueprintDescription}` |
| `console.*` / logger text | `console.error('Could not load the "…" catalog.', error)` |
| LLM prompt text | prompt/system-message strings sent to a model |
| test assertions and fixtures | anything in `*.test.ts` / `*.test.tsx` (also excluded from extraction) |
| SVG geometry | `d="M5.042 15.165a2.528 2.528 0 0 1-2.52 2.523…"` |
| brand and proper nouns standing alone | Slack, Notion, Linear, Spotify, Supabase, ZoomInfo, HomeAssistant, Confluence, Microsoft, Google, GitHub, Cloudflare, Workers, Jira, OpenAI, Anthropic |
| technical literals inside `<code>` `<kbd>` `<pre>` `<samp>` `<var>` | `<kbd>Esc</kbd>`, `<pre>` output blocks |
| `rel`, `sandbox`, `target`, `type`, `id`, `key`, `name` attributes | `rel="noopener noreferrer"` |

Two clarifications, because they look contradictory:

- A brand name **inside** a sentence you are wrapping stays in the sentence:
  `<Trans>Add credits to your AI Gateway</Trans>` is right. A brand name **alone** in a `<span>` is
  not a message.
- A technical literal inside a sentence you are wrapping becomes `<0>` content (E4). It is only
  "don't touch" when it stands alone.

### 7.2 Not in the residue report → not your job

The report is generated by a detector that sees JSX **text nodes** and JSX **string-literal
attributes**. It does not see:

- string literals inside JSX expressions — `{isProcessing ? 'Enabling...' : 'Always approve'}`
  (`src/components/AutoApproveConfirmDialog.tsx:80`), `{sk.note || 'Untitled link'}`
- template-literal attributes — ``title={`Remove “${removeOutput?.title || 'Untitled'}”?`}``
  (`src/routes/outputs.tsx:700`), ``aria-label={`Name for ${bindingName}`}``
  (`src/components/BlueprintBindingCard.tsx:68`)
- anything in a `.ts` file

**Leave all of it exactly as it is.** A later phase owns those. The only exception is when such a
string is inside the element you are rewriting and moving it is required to produce one message —
E9's `'Recipients must'` and E10's `"agents"`/`"people"` are inside their group and therefore in
scope.

---

## 8. Escalate, never guess

**Escalating is a SUCCESS.** An escalated entry costs the project two minutes of a senior model's
time. A wrong wrap ships broken copy or an untranslatable message that nobody notices for months.
You are not measured on how many you wrapped. You are measured on how many you got right plus how
accurately you flagged the rest. If you are not sure, escalate — that is the expected outcome for a
meaningful share of entries, and nobody will think less of the run for it.

### 8.1 Mandatory escalations — do not attempt these

| file | lines | why |
|---|---|---|
| `src/ChatInterface.tsx` | 3264–3265 | pluralizes by appending `"s"` to a **runtime** noun (`logKind`); not expressible in ICU without a copy decision |
| `src/AdminPage.tsx` | 753 | placeholder value contains a `\n` escape; template-literal equivalence not safe to assume |

Line 161 of `OutOfCreditsModal.tsx` is **not** escalated — it is an ordinary E3.

**SOLVED — no longer escalate these.** `src/components/billing/OutOfCreditsModal.tsx` 104–114 and
121–134 were listed here as mandatory escalations ("nested conditionals + two manual plurals + a
component placeholder in one paragraph"). Both are solved, in this repository, and 121–134 is now
the worked example in §6.2. Wrap them; do not spend an escalation re-deriving the answer.

What made them look unresolvable was reading each paragraph as one unit. Each is **two sentences**,
and once you split at the sentence boundary (§6.2) each conditional attaches to one sentence and the
difficulty disappears: sentence 1 owns the count plural, sentence 2 owns the `resetAt` conditional,
the second plural, and the `<ResetCountdown/>` placeholder. Neither needed a copy change or a
product decision. Rendered English is byte-identical in all four state combinations.

### 8.2 Escalate on sight

- You cannot tell whether the string is user-facing.
- The concatenation crosses a function or component boundary (part of the sentence is built
  somewhere else).
- The string is assembled by a helper (`buildLabel(...)`, `formatX(...)`) rather than written inline.
- Wrapping it would require changing behavior, control flow, types, or props.
- Wrapping it would require editing a file outside your shard.
- English pluralization is attached to a runtime value (`{noun}s`, `{word}(s)`).
- You cannot make the wrapped English render byte-identically to the current code.
- The typecheck, tests, or guard fail after your edit and you cannot see why in one look — revert
  that one edit and escalate it.

### 8.3 Return format

End your run with exactly this JSON object and nothing else after it:

```json
{
  "wrapped": 8,
  "escalated": [
    { "file": "packages/workshop-frontend/src/AdminPage.tsx", "line": 753,
      "reason": "mandatory: placeholder value contains a \\n escape" }
  ],
  "adjacentNotes": [
    { "file": "packages/workshop-frontend/src/ChatInterface.tsx", "line": 7112,
      "reason": "adjacent: wrapped the entry; the ternary on 7113-7115 builds an English sentence from template literals and is out of scope" }
  ]
}
```

- `wrapped` counts **residue entries** you resolved, not edits. Three entries sharing a `group` that
  you fixed with one `<Trans>` count as 3.
- Every entry in your shard must appear in `wrapped` or in `escalated`. `wrapped + escalated.length`
  must equal the number of residue entries assigned to you. Silently skipping one is the only
  outright failure mode. Do the arithmetic yourself before you report.
- `adjacentNotes` is **not** part of that sum. A §6.4-style adjacent observation is an entry you
  *did* wrap — it is already counted in `wrapped` — plus a heads-up about untouchable code beside it.
  Putting it in `escalated` would double-count it and break the invariant, which is why it has its
  own array. Keep the `adjacent:` prefix on the reason.
- `reason` is one sentence, concrete. Prefix `mandatory:` for §8.1 in `escalated`, `adjacent:` for
  §6.4-style notes in `adjacentNotes`.

In the example above: 9 entries were assigned, 8 were wrapped, 1 was escalated, `8 + 1 == 9`. One of
the 8 carries an adjacent note.

---

## 9. Verify before you report

Run all four from the worktree root. Copy-paste them exactly.

```bash
cd ~/repos/cloudflare-os-i18n-core

# 1. Typecheck the frontend (noEmit, safe to run while other agents work).
pnpm --filter @gadgets/workshop-frontend exec tsc -p tsconfig.json

# 2. Frontend test suite. Expect: Test Files 30 passed (30) · Tests 156 passed (156).
pnpm -C packages/workshop-frontend run test:run

# 3. Guard: your files must produce ZERO lines. Replace the names with your shard's files.
node scripts/check-unlocalized-strings.ts 2>/dev/null | grep -E 'AdminPage\.tsx|ShareModal\.tsx'
```

For step 3: the guard exits 1 while any file anywhere still has residue, so ignore its exit code and
read the grep output. Zero matching lines for your files (except any you escalated — those will
still be listed, which is correct) means you are done. `--dir` takes a directory, not a file, if you
prefer `node scripts/check-unlocalized-strings.ts --dir packages/workshop-frontend/src/components/billing`.

### 9.1 Exit codes are meaningless mid-wave — read the findings, not the code

This applies to **all** of the commands above and to `vp lint`, not just the guard. Several agents
run at once, and another agent's half-finished file makes `tsc` and `vp lint` exit non-zero — an
unbalanced JSX tag or a smart quote in someone else's file produces a wall of `TS1127`/`Invalid
Character` that has nothing to do with you. So:

- **Judge only findings that name a file in your shard.** Filter every command through your own file
  list, exactly as step 3 does:

  ```bash
  pnpm -C packages/workshop-frontend exec tsc 2>&1 | grep -E 'YourFile\.tsx|YourOther\.tsx'
  pnpm exec vp lint                 2>&1 | grep -E 'YourFile\.tsx|YourOther\.tsx'
  ```

  Zero lines from both = your edits are clean. Report the unfiltered totals too, and say plainly
  which files the remainder belongs to.
- **Never treat another file's failure as cover for one of your own.** "The build was already red"
  is not a passing result. If a finding names your file, it is yours, no matter how many other
  findings are on screen. Agents on this task have shipped broken files on exactly this
  rationalization — it is the most likely way for a defect to escape this pipeline.
- The test suite is the exception: it must be **green in full** (31 files / 157 tests as of
  2026-08-19; it was 30 / 156 while the extraction waves ran). Tests import
  across shards, so a red suite is everyone's problem. If it is red and none of the failures touch
  your files, say so explicitly and name the failing files rather than passing over it.
- Do not "fix" a finding in a file outside your shard, even an obvious one. Report it instead — the
  owning agent may be mid-edit, and two agents writing one file is how work gets lost.
- If you attribute a count to a file, **measure it**. Comparing two totals tells you a number
  changed, not which file changed it; per-file counts from one run are a ranking, not a delta. Say
  "I did not measure this" rather than inferring a cause.

Uncached full build — **the orchestrator runs this per wave; only run it if you are told to**:

```bash
pnpm exec vp run -r --no-cache build
```

> `pnpm run build --no-cache` is **WRONG**. The script forwards trailing flags to `tsc`, which dies
> with `error TS5023: Unknown compiler option '--no-cache'`. Use the `pnpm exec vp run -r --no-cache build`
> form above.

Do **not** run `pnpm run i18n:extract`. It rewrites `packages/i18n/locales/en.po`, which every agent
would collide on. The orchestrator runs it once after the wave.

If step 1 or 2 fails, fix your own edits. If you cannot, revert the single edit that caused it and
escalate that entry. Never weaken a test, never add `any`, never add `@ts-expect-error`, never add a
string to the allowlist.
