import { describe, expect, it } from 'vitest'
import { getPlaceholderRanges } from './resourceMatching'

function placeholderTexts(url: string): string[] {
  return getPlaceholderRanges(url).map(({ start, end }) => url.slice(start, end))
}

describe('getPlaceholderRanges', () => {
  it('ignores the literal share-link segment in a SharePoint URL', () => {
    expect(getPlaceholderRanges(
      'https://contoso.sharepoint.com/:l:/r/sites/Engineering/Lists/DEV%20Contoso%20OS%20Issues%20Log?e=AbC123',
    )).toEqual([])
  })

  it('finds a named parameter in a resource pattern', () => {
    const url = 'https://dash.cloudflare.com/:accountId/workers-and-pages/observability'
    expect(getPlaceholderRanges(url)).toEqual([
      { start: url.indexOf(':accountId'), end: url.indexOf(':accountId') + ':accountId'.length },
    ])
  })

  it('finds only the named parameter when a share-link segment is also present', () => {
    expect(placeholderTexts('https://contoso.sharepoint.com/:w:/r/sites/:siteName/Shared%20Documents'))
      .toEqual([':siteName'])
  })
})
