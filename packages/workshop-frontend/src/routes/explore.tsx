import { createFileRoute } from '@tanstack/react-router'
import { t } from '@lingui/core/macro'
import BlueprintsPage from '../BlueprintsPage'
import { useDocumentTitle } from '../useDocumentTitle'

export const Route = createFileRoute('/explore')({
  component: ExplorePage,
})

function ExplorePage() {
  useDocumentTitle(t`Explore`)

  return <BlueprintsPage />
}
