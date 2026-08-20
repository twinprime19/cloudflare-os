import { createFileRoute } from '@tanstack/react-router'
import { useState, useEffect, useMemo, useRef } from 'react'
import { DropdownMenu, useKumoToastManager } from '@cloudflare/kumo'
import { RpcStub } from 'capnweb'
import { useAuthenticatedApi } from '../AuthContext'
import {
  AdminApi,
  AiChatAuthorInfo,
  AiGatewayInfo,
  AiModelProvider,
  SUGGESTED_MODELS,
} from '@gadgets/workshop-shared/api'
import {
  Plus,
  Trash,
  Lightning,
  MagnifyingGlass,
  DotsThreeVertical,
  ShieldWarning,
  Star,
} from '@phosphor-icons/react'
import { t } from '@lingui/core/macro'
import { Trans } from '@lingui/react/macro'
import AddModelModal from '../AddModelModal'
import { useDocumentTitle } from '../useDocumentTitle'
import { MENU_CONTENT, MENU_ITEM, MENU_ITEM_DANGER } from '../components/menuStyles'

export const Route = createFileRoute('/providers')({ component: ProvidersPage })

// ─── constants ────────────────────────────────────────────────────────────────

const PROVIDER_ORDER = Object.keys(SUGGESTED_MODELS) as AiModelProvider[]

const PRIMARY_BTN =
  'press inline-flex h-9 shrink-0 cursor-pointer items-center justify-center gap-1.5 rounded-lg bg-kumo-brand px-3.5 text-[13px] font-medium tracking-[-0.25px] text-white transition-colors hover:bg-kumo-brand-hover disabled:cursor-default disabled:opacity-50'

// ─── model row ─────────────────────────────────────────────────────────────────

// Rows mirror the Blueprints list: a clickable row (here, clicking sets/clears the model every
// user's new chats start on) plus a kebab for the rest. The whole row is the primary affordance,
// so it shows a pointer.
function ModelRow({
  model,
  isDefault,
  isBuiltIn,
  onDelete,
  onSetDefault,
}: {
  model: AiChatAuthorInfo
  isDefault: boolean
  isBuiltIn: boolean
  onDelete: () => void
  onSetDefault: () => void
}) {
  return (
    <div
      role="button"
      tabIndex={0}
      onClick={onSetDefault}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault()
          onSetDefault()
        }
      }}
      title={isDefault ? t`Default model. Click to clear` : t`Click to set as default model`}
      className="group flex cursor-pointer items-center gap-3 rounded-lg px-3 py-2.5 transition-colors duration-150 ease-out hover:bg-kumo-tint"
    >
      {/* Neutral monogram — matches the sidebar/workspaces treatment */}
      <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-kumo-fill text-[12px] font-medium text-kumo-subtle">
        {model.name[0]?.toUpperCase()}
      </div>

      {/* Info */}
      <div className="min-w-0 flex-1">
        <div className="flex items-center gap-2">
          <span className="truncate text-sm font-medium tracking-[-0.25px] text-kumo-default">
            {model.name}
          </span>
          {isBuiltIn && (
            <span className="shrink-0 rounded-full bg-kumo-tint px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-[0.4px] text-kumo-subtle">
              built-in
            </span>
          )}
          {isDefault && (
            <span className="inline-flex shrink-0 items-center gap-1 rounded-full bg-[rgba(255,72,1,0.10)] px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-[0.4px] text-kumo-brand">
              <Star size={9} weight="fill" />
              <Trans>Default</Trans>
            </span>
          )}
        </div>
        <span className="mt-0.5 block truncate font-mono text-[12px] tracking-[-0.1px] text-kumo-inactive">
          {model.id}
        </span>
      </div>

      {/* Actions */}
      <div onClick={(e) => { e.stopPropagation() }}>
        <DropdownMenu>
          <DropdownMenu.Trigger
            render={
              <button
                aria-label={t`Provider actions`}
                className="cursor-pointer rounded-md p-1.5 text-kumo-subtle transition-colors hover:bg-kumo-fill hover:text-kumo-default focus:opacity-100 sm:opacity-0 sm:group-hover:opacity-100"
              >
                <DotsThreeVertical size={16} />
              </button>
            }
          />
          <DropdownMenu.Content className={MENU_CONTENT}>
            <DropdownMenu.Item onClick={onSetDefault} className={MENU_ITEM}>
              <Star size={13} className="mr-2" weight={isDefault ? 'fill' : 'regular'} />
              {isDefault ? t`Clear default model` : t`Set as default`}
            </DropdownMenu.Item>
            {!isBuiltIn && (
              <DropdownMenu.Item variant="danger" onClick={onDelete} className={MENU_ITEM_DANGER}>
                <Trash size={13} className="mr-2" />
                <Trans>Delete provider</Trans>
              </DropdownMenu.Item>
            )}
          </DropdownMenu.Content>
        </DropdownMenu>
      </div>
    </div>
  )
}

// ─── notice ────────────────────────────────────────────────────────────────────

function Notice({ children }: { children: React.ReactNode }) {
  return (
    <div className="flex items-start gap-3 rounded-xl border border-kumo-line bg-kumo-tint px-4 py-3 text-[13px] leading-[18px] tracking-[-0.25px] text-kumo-subtle">
      {children}
    </div>
  )
}

// ─── main page ────────────────────────────────────────────────────────────────

function ProvidersPage() {
  useDocumentTitle('AI Providers')

  const { authenticatedApi, isAdmin } = useAuthenticatedApi()
  const toasts = useKumoToastManager()
  // The admin capability (minted once via getAdminApi; null until loaded / for non-admins). Wrapped
  // in an object so useState doesn't treat the (callable) RPC stub as a state updater function.
  const [admin, setAdmin] = useState<{ api: RpcStub<AdminApi> } | null>(null)
  const [sharedModels, setSharedModels] = useState<AiChatAuthorInfo[]>([])
  const [defaultModelId, setDefaultModelId] = useState<string | null>(null)
  const [aiConfig, setAiConfig] = useState<AiGatewayInfo | null>(null)
  const [search, setSearch] = useState('')
  const [loading, setLoading] = useState(true)
  const [loadError, setLoadError] = useState(false)
  const [sheetOpen, setSheetOpen] = useState(false)
  const [deletingId, setDeletingId] = useState<string | null>(null)

  // Mint the admin capability once (the access check happens server-side) and load the catalog.
  useEffect(() => {
    if (!isAdmin) {
      setLoading(false)
      return
    }
    let cancelled = false
    let stub: RpcStub<AdminApi> | null = null
    ;(async () => {
      try {
        const api = await authenticatedApi.getAdminApi()
        if (cancelled) {
          api?.[Symbol.dispose]?.()
          return
        }
        if (!api) {
          setLoadError(true)
          return
        }
        stub = api
        setAdmin({ api })
        const [settings, cfg] = await Promise.all([api.getSettings(), authenticatedApi.getAiConfig()])
        if (cancelled) return
        setSharedModels(settings.sharedModels)
        setDefaultModelId(settings.defaultModelId)
        setAiConfig(cfg)
      } catch (err) {
        if (!cancelled) {
          console.error('Failed to load providers:', err)
          setLoadError(true)
        }
      } finally {
        if (!cancelled) setLoading(false)
      }
    })()
    return () => {
      cancelled = true
      stub?.[Symbol.dispose]?.()
    }
  }, [isAdmin, authenticatedApi])

  // Re-read the catalog after a write. The gateway config is deployment-fixed, so it is not refetched.
  const refreshModels = async () => {
    if (!admin) return
    const settings = await admin.api.getSettings()
    setSharedModels(settings.sharedModels)
    setDefaultModelId(settings.defaultModelId)
  }

  const gatewayMode = aiConfig?.enabled === true

  // The gateway's built-in models, which every user already has and no admin can remove.
  const builtInModels = useMemo<AiChatAuthorInfo[]>(() => {
    if (!aiConfig?.enabled) return []
    const enabled = new Set(aiConfig.enabledProviders)
    return PROVIDER_ORDER.filter((provider) => enabled.has(provider)).flatMap((provider) =>
      Object.entries(SUGGESTED_MODELS[provider]).map(([id, model]) => ({
        type: 'agent' as const,
        id,
        name: model.name,
      }))
    )
  }, [aiConfig])

  // What the deployment offers, in the order every user sees it: built-ins, then shared models. A
  // shared entry that duplicates a built-in is unreachable — the gateway's model is the one chats
  // resolve — so it is left out here too.
  const rows = useMemo(() => {
    const builtInIds = new Set(builtInModels.map((model) => model.id))
    return [
      ...builtInModels.map((model) => ({ model, builtIn: true })),
      ...sharedModels
        .filter((model) => !builtInIds.has(model.id))
        .map((model) => ({ model, builtIn: false })),
    ]
  }, [builtInModels, sharedModels])

  const handleDelete = async (model: AiChatAuthorInfo) => {
    if (!admin) return
    if (!confirm(`Delete "${model.name}"? This cannot be undone.`)) return
    setDeletingId(model.id)
    try {
      await admin.api.deleteSharedModel(model.id)
      await refreshModels()
    } catch (err) {
      console.error('Failed to delete model:', err)
      toasts.add({ title: 'Failed to delete provider', variant: 'error' })
    } finally {
      setDeletingId(null)
    }
  }

  // Overlapping setDefaultModel calls have no ordering guarantee, so ignore clicks while one is
  // in flight.
  const defaultInFlight = useRef(false)
  const handleSetDefault = async (modelId: string) => {
    if (!admin || defaultInFlight.current) return
    defaultInFlight.current = true
    const previous = defaultModelId
    const next = defaultModelId === modelId ? null : modelId
    setDefaultModelId(next)
    try {
      await admin.api.setDefaultModel(next)
    } catch (err) {
      console.error('Failed to set default model:', err)
      setDefaultModelId(previous) // revert
      toasts.add({ title: 'Failed to update default model', variant: 'error' })
    } finally {
      defaultInFlight.current = false
    }
  }

  const filtered = rows.filter(({ model }) => {
    if (!search) return true
    const q = search.toLowerCase()
    return model.name.toLowerCase().includes(q) || model.id.toLowerCase().includes(q)
  })

  if (!isAdmin) {
    return (
      <div className="max-w-2xl mx-auto px-4 sm:px-6 py-16 text-center">
        <ShieldWarning size={32} className="mx-auto text-kumo-subtle mb-3" />
        <p className="text-sm text-kumo-default"><Trans>You don't have access to this page.</Trans></p>
      </div>
    )
  }

  return (
    <div className="mx-auto flex h-full w-full max-w-4xl flex-col px-6 sm:px-10">
      <header className="flex items-end justify-between gap-4 px-3 pb-3 pt-10">
        <div className="min-w-0">
          <h1 className="text-2xl font-semibold tracking-tight text-kumo-default"><Trans>AI providers</Trans></h1>
          <p className="mt-1 text-[13px] leading-[18px] tracking-[-0.25px] text-kumo-subtle">
            <Trans>Manage the AI models everyone on this deployment can use.</Trans>
          </p>
        </div>
        <button type="button" onClick={() => setSheetOpen(true)} disabled={!admin} className={PRIMARY_BTN}>
          <Plus size={14} weight="bold" />
          <Trans>Add provider</Trans>
        </button>
      </header>

      {/* Search — hidden when the deployment has no models */}
      {!loading && !loadError && rows.length > 0 && (
        <div className="mb-3 px-3">
          <div className="relative">
            <MagnifyingGlass size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-kumo-inactive" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder={t`Search providers…`}
              className="h-9 w-full rounded-lg border border-kumo-line bg-kumo-base pl-9 pr-4 text-[13px] tracking-[-0.25px] text-kumo-default placeholder:text-kumo-inactive transition-[border-color,box-shadow] duration-150 ease-out focus:border-kumo-ring focus:outline-none focus:ring-[3px] focus:ring-kumo-ring/15"
            />
          </div>
        </div>
      )}

      <div className="chat-panel flex min-h-0 flex-1 flex-col gap-0.5 overflow-y-auto pt-1 pb-16">
        {/* Notices */}
        {gatewayMode && !loading && !loadError && (
          <div className="flex flex-col gap-2.5 px-3 pb-2">
            <Notice>
              <Lightning size={15} className="mt-px shrink-0 text-kumo-brand" />
              <span>
                <Trans>
                  <strong className="font-medium text-kumo-default">AI Gateway mode:</strong> built-in
                  models are managed by your deployment. You can still add custom models with your own
                  API tokens.
                </Trans>
              </span>
            </Notice>
          </div>
        )}

        {/* Model list */}
        {loading ? (
          <div className="flex flex-col gap-0.5 px-3">
            {[1, 2, 3].map((i) => (
              <div key={i} className="h-[56px] animate-pulse rounded-xl bg-kumo-elevated" />
            ))}
          </div>
        ) : loadError ? (
          <div className="py-12 text-center text-sm">
            <p className="text-kumo-danger"><Trans>Something went wrong loading your providers.</Trans></p>
            <button type="button" onClick={() => window.location.reload()} className="mt-1 cursor-pointer text-kumo-brand underline">
              <Trans>Try again</Trans>
            </button>
          </div>
        ) : rows.length === 0 ? (
          <div className="flex flex-col items-center gap-3 px-3 py-16 text-center">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-kumo-fill text-kumo-subtle">
              <Lightning size={18} />
            </div>
            <div>
              <p className="text-sm font-medium text-kumo-default"><Trans>No AI providers yet</Trans></p>
              <p className="mt-1 text-[13px] leading-[18px] text-kumo-subtle">
                <Trans>Add a provider to start building workspaces with AI.</Trans>
              </p>
            </div>
            <button type="button" onClick={() => setSheetOpen(true)} disabled={!admin} className={PRIMARY_BTN}>
              <Plus size={14} weight="bold" />
              <Trans>Add your first provider</Trans>
            </button>
          </div>
        ) : filtered.length === 0 ? (
          <div className="py-12 text-center text-sm text-kumo-inactive"><Trans>No providers found</Trans></div>
        ) : (
          filtered.map(({ model, builtIn }) => (
            <div
              key={model.id}
              className={deletingId === model.id ? 'pointer-events-none opacity-50' : ''}
            >
              <ModelRow
                model={model}
                isDefault={defaultModelId === model.id}
                isBuiltIn={builtIn}
                onDelete={() => handleDelete(model)}
                onSetDefault={() => handleSetDefault(model.id)}
              />
            </div>
          ))
        )}
      </div>

      {/* Add model dialog — a model added here is offered to every user of the deployment. */}
      {admin && (
        <AddModelModal
          visible={sheetOpen}
          onCancel={() => setSheetOpen(false)}
          onSuccess={() => {
            setSheetOpen(false)
            refreshModels().catch((err) => {
              console.error('Failed to reload providers:', err)
              toasts.add({ title: 'Failed to reload providers', variant: 'error' })
            })
          }}
          onSubmit={(profile, config) => admin.api.addSharedModel(profile, config)}
          aiConfig={aiConfig}
        />
      )}
    </div>
  )
}
