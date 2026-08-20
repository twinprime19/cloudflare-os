import { setupI18n, type I18n } from '@lingui/core'
import { describe, expect, it } from 'vitest'
import { SOURCE_LOCALE, i18n } from '@gadgets/i18n'

// The message `AutoApproveConfirmDialog.tsx` wraps, and the value its placeholder must be replaced
// with. Chosen because it is interpolated: a plain message survives the failure this file guards.
const INTERPOLATED_MESSAGE = 'Always approve “{actionLabel}”?'
const INTERPOLATED_RESULT = 'Always approve “Delete file”?'
const PLACEHOLDER = 'actionLabel'

/** Node's `process`, typed here because the app's tsconfig deliberately has no `node` types. */
const nodeProcess = (globalThis as unknown as {
  process: { env: Record<string, string | undefined> }
}).process

/**
 * A `@lingui/core` instance built the way a production bundle builds one.
 *
 * Its constructor registers the runtime ICU compiler only when `NODE_ENV !== 'production'`, so in a
 * production build every message that resolves to a raw ICU string is returned unparsed and
 * uninterpolated -- and a raw ICU string is exactly what `descriptorFields: 'message'` leaves
 * inline at each call site. Vitest runs as `NODE_ENV=test`, where that compiler IS registered, so
 * the defect is invisible unless a test constructs the instance under the other branch on purpose.
 */
function productionI18n(): I18n {
  const previous = nodeProcess.env.NODE_ENV
  nodeProcess.env.NODE_ENV = 'production'
  try {
    return setupI18n()
  } finally {
    if (previous === undefined) delete nodeProcess.env.NODE_ENV
    else nodeProcess.env.NODE_ENV = previous
  }
}

/**
 * Rebuilds the source ICU text a compiled catalog entry represents. Catalog keys are opaque
 * hashes -- not the message text -- so the only way to find "the entry for this exact message" is
 * to reconstruct each entry's text and compare it, rather than sniff the token shape for a
 * placeholder name. A compiled entry is either a plain string (no placeholders) or an array mixing
 * literal string chunks with `[name, ...]` argument tokens, which this puts back as `{name}`. Good
 * enough for the plain-interpolation messages this file compares against; an unrecognised token
 * shape (e.g. plural/select forms) returns `null` rather than guess, so it simply never matches.
 */
function reconstructSource(compiled: unknown): string | null {
  if (typeof compiled === 'string') return compiled
  if (!Array.isArray(compiled)) return null
  let text = ''
  for (const token of compiled) {
    if (typeof token === 'string') text += token
    else if (Array.isArray(token) && typeof token[0] === 'string') text += `{${token[0]}}`
    else return null
  }
  return text
}

/**
 * The id of the catalog entry for {@link INTERPOLATED_MESSAGE}, found by exact source text rather
 * than "the first entry with an `actionLabel` placeholder" -- more than one message carries that
 * placeholder now (`AutoApproveConfirmDialog.tsx`'s own always-approve confirmation gained a second
 * one), and matching on token shape alone silently retargeted this test at the wrong entry.
 * Asserting the match set is exactly one keeps that failure loud if a third message compiling to
 * this exact text ever appears.
 */
function interpolatedMessageId(): string {
  const matches = Object.entries(i18n.messages).filter(
    ([, compiled]) => reconstructSource(compiled) === INTERPOLATED_MESSAGE,
  )
  expect(matches, `exactly one catalog entry should compile ${JSON.stringify(INTERPOLATED_MESSAGE)}`)
    .toHaveLength(1)
  return matches[0][0]
}

describe('source catalog', () => {
  // Fails if the source catalog is fetched rather than statically imported: nothing has rendered
  // yet at this point, so a provider-driven load would not have run.
  it('is compiled, loaded and active from the moment the package is imported', () => {
    expect(i18n.locale).toBe(SOURCE_LOCALE)
    expect(Object.keys(i18n.messages).length).toBeGreaterThan(0)
  })
})

describe('message resolution under production semantics', () => {
  it('substitutes placeholders when the compiled catalog is loaded', () => {
    const production = productionI18n()
    production.loadAndActivate({ locale: SOURCE_LOCALE, messages: i18n.messages })

    const rendered = production._({
      id: interpolatedMessageId(),
      message: INTERPOLATED_MESSAGE,
      values: { actionLabel: 'Delete file' },
    })

    expect(rendered).toBe(INTERPOLATED_RESULT)
  })

  // Characterizes the failure the static import exists to prevent: with no compiled catalog, the
  // inline raw-ICU fallback is returned verbatim, braces and all. This is `@lingui/core` behaviour
  // rather than ours -- it is pinned here so anyone tempted to make the source catalog lazy again
  // can see exactly what that costs.
  it('renders the placeholder literally when no compiled catalog is loaded', () => {
    const production = productionI18n()
    production.loadAndActivate({ locale: SOURCE_LOCALE, messages: {} })

    const rendered = production._({
      id: 'an-id-that-is-not-in-the-catalog',
      message: INTERPOLATED_MESSAGE,
      values: { actionLabel: 'Delete file' },
    })

    expect(rendered).toBe(INTERPOLATED_MESSAGE)
    expect(rendered).toContain(`{${PLACEHOLDER}}`)
  })
})
