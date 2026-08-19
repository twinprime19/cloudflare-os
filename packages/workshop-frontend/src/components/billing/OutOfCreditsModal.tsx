import { useCallback, useEffect, useState } from 'react'
import { CloudflareUsageInfo, CloudflareAccountOption } from '@gadgets/workshop-shared/api'
import { Dialog, Button, Loader, useKumoToastManager } from '@cloudflare/kumo'
import { CloudWarning, Lightning } from '@phosphor-icons/react'
import { Plural, Trans } from '@lingui/react/macro'
import { useOptionalAuthenticatedApi } from '../../AuthContext'
import { buildAddCreditsUrl } from './creditsUrl'
import ResetCountdown from './ResetCountdown'

interface OutOfCreditsModalProps {
  open: boolean
  onClose: () => void
}

/**
 * Modal shown when a user has exhausted their free daily allowance. Guides them to connect their
 * Cloudflare account (if not connected), pick which account to bill (if they have several), or top
 * up credits in the Cloudflare dashboard (if connected but low balance).
 */
export default function OutOfCreditsModal({ open, onClose }: OutOfCreditsModalProps) {
  const auth = useOptionalAuthenticatedApi()
  const toasts = useKumoToastManager()
  const [usage, setUsage] = useState<CloudflareUsageInfo | null>(null)
  const [connecting, setConnecting] = useState(false)
  const [accounts, setAccounts] = useState<CloudflareAccountOption[] | null>(null)
  const [selecting, setSelecting] = useState<string | null>(null)

  const refresh = useCallback(() => {
    if (!auth) return
    auth.authenticatedApi.getCloudflareUsage()
      .then((u: CloudflareUsageInfo) => setUsage(u))
      .catch(() => {})
  }, [auth])

  useEffect(() => {
    if (!open || !auth) return
    setUsage(null)
    setAccounts(null)
    refresh()
    // Re-check when the tab regains focus, so returning from the "Connect Cloudflare" OAuth pop-up
    // updates the modal (connected state / balance / account list) without reopening it.
    const onFocus = () => {
      setAccounts(null)
      refresh()
    }
    window.addEventListener('focus', onFocus)
    return () => window.removeEventListener('focus', onFocus)
  }, [open, auth, refresh])

  // Load the account list when the server says the user must pick one.
  useEffect(() => {
    if (!auth) return
    if (usage?.connected && usage.needsAccountSelection && accounts === null) {
      auth.authenticatedApi.listCloudflareAccounts()
        .then((list: CloudflareAccountOption[]) => setAccounts(list))
        .catch(() => setAccounts([]))
    }
  }, [auth, usage, accounts])

  const connect = async () => {
    if (!auth) return
    setConnecting(true)
    try {
      const { url } = await auth.authenticatedApi.connectAccount('cloudflare')
      window.open(url, '_blank', 'noopener,noreferrer')
    } catch {
      // ignore
    } finally {
      setConnecting(false)
    }
  }

  const selectAccount = async (accountId: string) => {
    if (!auth) return
    setSelecting(accountId)
    try {
      await auth.authenticatedApi.selectCloudflareAccount(accountId)
      setAccounts(null)
      refresh()
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to select account'
      toasts.add({ title: msg, variant: 'error' })
    } finally {
      setSelecting(null)
    }
  }

  const connected = usage?.connected ?? false
  const needsSelection = connected && (usage?.needsAccountSelection ?? false)

  return (
    <Dialog.Root open={open} onOpenChange={(o) => { if (!o) onClose() }}>
      <Dialog className="p-6 sm:w-[560px]" size="base">
        <Dialog.Title className="text-lg font-semibold mb-2 flex items-center gap-2">
          <CloudWarning size={22} weight="bold" className="text-kumo-warning" />
          <Trans>You've reached your free usage limit</Trans>
        </Dialog.Title>

        {usage === null ? (
          <div className="flex justify-center py-8"><Loader size="base" /></div>
        ) : (
          <div className="space-y-4">
            {!connected ? (
              <p className="text-sm text-kumo-subtle">
                <Trans>
                  You've used all{' '}
                  <Plural
                    value={usage.dailyLimit}
                    one="# of your free request"
                    other="# of your free requests"
                  />{' '}
                  for today.
                </Trans>{' '}
                {usage.resetAt ? (
                  <Trans>
                    Connect your Cloudflare account to keep building now — usage beyond the free tier
                    is billed to your own Cloudflare AI Gateway credits — or wait: your free{' '}
                    <Plural value={usage.dailyLimit} one="request resets" other="requests reset" /> at
                    00:00 UTC, in <ResetCountdown resetAt={usage.resetAt} onElapsed={refresh} />.
                  </Trans>
                ) : (
                  <Trans>
                    Connect your Cloudflare account to keep building now — usage beyond the free tier
                    is billed to your own Cloudflare AI Gateway credits.
                  </Trans>
                )}
              </p>
            ) : needsSelection ? (
              <p className="text-sm text-kumo-subtle">
                <Trans>Your Cloudflare connection has access to multiple accounts. Choose which one's AI
                Gateway credits should be billed for usage beyond the free tier.</Trans>
              </p>
            ) : (
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
            )}

            {needsSelection && (
              <div className="flex flex-col gap-2">
                {accounts === null ? (
                  <p className="text-sm text-kumo-subtle"><Trans>Loading accounts…</Trans></p>
                ) : accounts.length === 0 ? (
                  <p className="text-sm text-kumo-subtle"><Trans>No accounts available on this connection.</Trans></p>
                ) : (
                  accounts.map((a) => (
                    <Button
                      key={a.accountId}
                      variant="secondary"
                      className="justify-start"
                      onClick={() => selectAccount(a.accountId)}
                      loading={selecting === a.accountId}
                      disabled={selecting !== null}
                    >
                      {a.accountName}
                    </Button>
                  ))
                )}
              </div>
            )}

            <p className="text-sm text-kumo-subtle">
              <Trans>
                Learn more about{' '}
                <a
                  href="https://developers.cloudflare.com/ai-gateway/features/unified-billing/"
                  target="_blank"
                  rel="noreferrer"
                  className="underline"
                >
                  AI Gateway unified billing
                </a>
                .
              </Trans>
            </p>

            <div className="flex items-center justify-end gap-2 pt-2">
              {!connected ? (
                <>
                  <Button variant="secondary" onClick={onClose}><Trans>Maybe later</Trans></Button>
                  <Button variant="primary" onClick={connect} loading={connecting}>
                    <Lightning size={16} weight="bold" />
                    <Trans>Connect Cloudflare</Trans>
                  </Button>
                </>
              ) : needsSelection ? (
                <Button variant="secondary" onClick={onClose}><Trans>Close</Trans></Button>
              ) : (
                <>
                  <Button variant="secondary" onClick={onClose}><Trans>Close</Trans></Button>
                  <Button
                    variant="primary"
                    onClick={() => window.open(buildAddCreditsUrl(usage.accountId), '_blank')}
                  >
                    <Trans>Add credits in Cloudflare</Trans>
                  </Button>
                </>
              )}
            </div>
          </div>
        )}
      </Dialog>
    </Dialog.Root>
  )
}
