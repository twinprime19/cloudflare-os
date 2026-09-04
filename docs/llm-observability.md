# LLM observability (Langfuse)

Per-user and per-model usage analytics for the platform: who uses it, how much, how recently, on
which models, at what cost. Data lands in a self-hosted Langfuse instance. Nothing in this repo
talks to Langfuse directly.

## Architecture

```
app (workshop-backend) ──cf-aig-metadata header──▶ Cloudflare AI Gateway (platform gateway)
                                                        │  OpenTelemetry exporter (gateway config)
                                                        ▼
                                        https://<langfuse-host>/api/public/otel/v1/traces
```

- Every LLM call routed through the platform gateway becomes one Langfuse trace with model,
  input/output tokens, cost, latency, and the request metadata.
- Export is gateway-side and asynchronous: no Worker code, no added request latency.
- Prompt and completion bodies are included in the export. The destination is Contoso's own host, so
  there is no third-party egress, but treat Langfuse logins as sensitive.

## Where the pieces live

| Piece | Location |
|---|---|
| Request metadata | `buildMetadata()` in `packages/workshop-backend/src/ai-models.ts` |
| Exporter config | Cloudflare dashboard → AI Gateway → the platform gateway → Observability → OpenTelemetry |
| Langfuse stack | Docker host operated by Contoso (compose stack, not in this repo) |
| Dashboard definition | `docs/langfuse/user-usage-dashboard.json` |

Exporter settings: endpoint `https://<langfuse-host>/api/public/otel/v1/traces`, header
`Authorization: Basic base64(pk-lf-…:sk-lf-…)` using a Langfuse project API key pair.

## User attribution

The gateway forwards `cf-aig-metadata` keys verbatim as span attributes and keeps at most five
entries. Langfuse promotes only the `user.id` attribute to its first-class user field, so the
metadata carries the acting user's identifier under exactly that key. On gatekeeper sign-in
(Microsoft and other OAuth providers) the identifier is the verified login email.

- Gadget-driven calls are credited to the gadget's owner and additionally flagged `automated`.
- Hook-triggered gadget sessions are credited to the gadget owner, not the external caller.
- The other metadata keys are `source`, `gadgetId`, and `chatId`.
- Traces recorded before the `user.id` key shipped (2026-09-04) carry the identifier only inside
  trace metadata and do not appear on the Users page. They cannot be backfilled.
- In the Cloudflare AI Gateway logs UI, filter on metadata key `user.id`.

## Dashboards

- **Users page** (Langfuse left nav): every user with trace, token, and cost totals plus last
  activity. Sort by traces or tokens for "who uses the most".
- **User usage dashboard**: import `docs/langfuse/user-usage-dashboard.json` by dragging the file
  onto the Dashboards page. Requires Langfuse v3.212.0 or newer. Tiles: active users (distinct),
  requests, cost, active users over time, requests over time by user, top 20 users by requests,
  tokens, and cost, and a per-user activity table. All tiles follow the dashboard date range.
- Sanity check after import: the Active users tile should equal the row count of the User
  activity table over the same range.
- Interactive versus gadget-driven split: add a widget filter on column `metadata` in the widget
  editor and pick the key the editor suggests for the `automated` attribute.
- Keep staging and production gateways on separate Langfuse projects (each exporter carries its
  own key pair) so test traffic never inflates user counts.

Scripted access uses the Metrics API with a project key pair, for example distinct active users
and per-user totals over a window:

```json
{"view":"traces",
 "metrics":[{"measure":"count","aggregation":"count"},
            {"measure":"totalCost","aggregation":"sum"},
            {"measure":"totalTokens","aggregation":"sum"}],
 "dimensions":[{"field":"userId"}],
 "fromTimestamp":"<ISO start>","toTimestamp":"<ISO end>"}
```

Send it URL-encoded as the `query` parameter of `GET /api/public/metrics`. The row count is the
number of active users, subject to the API's row limit.

## Operations

- **Key rotation**: create a new API key pair in the Langfuse project, replace the Basic auth
  header on the gateway exporter, then delete the old pair.
- **Retention and backups** are owned on the Docker host (Postgres, ClickHouse, and object storage
  volumes).
- **Upgrades**: bump the pinned Langfuse image tag and restart the stack; migrations run on boot.

## Known gaps

- Only platform-gateway traffic is exported. A user who logs in but never sends an LLM request is
  invisible.
- Enabling `ENABLE_CLOUDFLARE_LIMITS` lets self-funded users route through their own AI Gateway.
  Those calls bypass the platform exporter and vanish from Langfuse. Do not enable it without
  deciding how usage tracking should follow.
- Deployments without a platform gateway (direct provider mode) send no metadata and export
  nothing.
- Cost in Langfuse is the gateway's estimate, suitable for analytics. Authoritative per-request
  cost is still read from the gateway log by the app (see `docs/ai-gateway-billing.md`).
