# Microsoft 365 gatekeeper

Microsoft Entra ID (Azure AD) integration for Gadgets. It serves two purposes:

- **Sign-in:** when `microsoft` is in the deployment's `AUTH_GATEKEEPERS` allowlist, "Continue with
  Microsoft" appears on the login page. Sign-in requests only the identity scopes (`openid`,
  `profile`, `email`, `User.Read`) to read the account's **verified email**, which becomes the
  user's identity. The sign-in grant is transient (discarded right after the email is read).
- **Connections:** when a user connects Microsoft, `Mail.ReadWrite` (plus `offline_access`) is
  requested for the Outlook mailbox at `https://outlook.office.com/mail/`, so gadgets can read,
  organize, and draft replies to mail on the user's behalf.

A single Entra app registration is used for both, and it is pinned to **one tenant** — the worker
needs `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` before it will answer anything.

## Identity policy

Only directory **members** on one of the tenant's **verified domains** can sign in. B2B guests
(`userType` of `Guest`, `#EXT#` accounts, or addresses on unverified domains) are rejected by
design: their addresses are owned by another tenant, so this tenant has not proven the email
belongs to the person signing in. The address is lowercased and becomes the account key, so signing
in with Microsoft resolves to the same Workshop account as any other gatekeeper that verifies the
same address.

## Setting up the Entra app registration

### Step 1: Register the application

1. Go to the [Microsoft Entra admin center](https://entra.microsoft.com) and sign in as an
   administrator of the tenant whose users should have access.
2. Go to **Applications** > **App registrations** > **New registration**.
3. **Name**: anything (e.g. "Cloudflare OS").
4. **Supported account types**: **Accounts in this organizational directory only (single tenant)**.
   The gatekeeper builds every OAuth URL from `TENANT_ID`, and only tenant members can sign in
   anyway, so a multi-tenant registration would only widen the app's exposure.
5. **Redirect URI**: platform **Web** (do not use "Single-page application"), value
   `${PUBLIC_BASE_URL}/gatekeeper/microsoft/oauth` — for local dev that is
   `http://localhost:8787/gatekeeper/microsoft/oauth` (no trailing slash, http not https). Choosing
   "Single-page application" instead causes sign-in to fail at code redemption with `AADSTS9002325`
   (the gatekeeper redeems the code itself using the client secret; SPA platform mandates browser
   PKCE, which this server-side flow does not use).
6. Click **Register**, then copy the **Application (client) ID** and **Directory (tenant) ID** from
   the Overview page.

### Step 2: Grant the API permissions

Under **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**,
add:

| Permission | Why |
| --- | --- |
| `openid`, `profile`, `email` | sign-in and the id token claims the email is read from |
| `offline_access` | refresh tokens, so a connected mailbox survives past the first hour |
| `User.Read` | the signed-in user's profile, and the tenant's verified domain list |
| `Mail.ReadWrite` | read messages and folders, flip read state, move messages, draft replies |

Then click **Grant admin consent for &lt;tenant&gt;**. Many tenants block user consent; without
admin consent those users hit `AADSTS65001` ("The user or administrator has not consented") on
their first sign-in and can never get past it.

### Step 3: Create a client secret — and diarize its expiry

Under **Certificates & secrets** > **Client secrets** > **New client secret**:

1. Pick an expiry. The admin center caps new secrets at **24 months** and commonly defaults to
   **180 days**. Longer means fewer rotations; shorter limits how long a leaked secret is useful.
   Pick deliberately — this is not a value to accept by reflex.
2. Click **Add** and copy the secret **Value** immediately. It is displayed once, and only the
   Secret ID is retrievable afterwards.
3. **Record the expiry date** somewhere your team will see it, with a reminder at least two weeks
   ahead.

> **The expiry is an outage, not a warning.** When the secret expires, Entra rejects every token
> exchange with `AADSTS7000222`. That error is only visible inside the OAuth pop-up: the main window
> is still waiting on the login attempt, so it spins forever. If the deployment also sets
> `DISABLE_PASSWORD_AUTH=true` and Microsoft is the only allowlisted sign-in gatekeeper, **nobody
> can log in at all** — including the administrator who would fix it. Rotation is an operational
> requirement here, not hygiene.

### Step 4: Rotation runbook

Rotate *before* the expiry date, never after. Entra allows more than one active secret, so there is
no gap:

1. **Create** a second client secret on the same app registration (Certificates & secrets > New
   client secret) and copy its value.
2. **Update the deployment** to the new value — reinstall/reconfigure the gatekeeper with the new
   `CLIENT_SECRET` (or `wrangler secret put CLIENT_SECRET` for a self-managed deploy), then confirm
   a fresh sign-in and a connected mailbox both still work. Existing sessions are unaffected; only
   token exchanges use the secret.
3. **Delete** the old secret once the new one is proven. Leaving it in place keeps a second valid
   credential alive for no reason.
4. Record the new expiry date and reset the reminder.

If the secret has already expired, the same steps apply — create a new secret and update the
deployment. Nothing else needs re-consenting, and connected mailboxes recover on their next refresh.

### Step 5: Configure the worker

The gatekeeper reads three vars: `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`. Deployments supply them
as secrets (the deploy wizard asks for all three; see `deploy-inputs.json`).

For local dev, put the shell vars in the gitignored root `.dev.vars` (or export them);
`run-dev-server.ts` maps them into the worker:

```
MICROSOFT_CLIENT_ID=<Application (client) ID>
MICROSOFT_CLIENT_SECRET=<client secret value>
MICROSOFT_TENANT_ID=<Directory (tenant) ID>
```

`MICROSOFT_TENANT_ID` seeds the worker's `TENANT_ID`. If it is missing while the client credentials
are present, the dev server prints a warning; the OAuth page then renders "Microsoft Gatekeeper Not
Configured" and connection attempts fail the same way.

### Step 6: (Optional) Enable Microsoft sign-in

To offer "Continue with Microsoft" on the login page, add `microsoft` to the deployment's
`AUTH_GATEKEEPERS` allowlist (e.g. in the root `.dev.vars`):

```
AUTH_GATEKEEPERS=microsoft
```

See [docs/oauth-signin.md](../../docs/oauth-signin.md) for how the allowlist and
`DISABLE_PASSWORD_AUTH` interact.

## Manual verification

Everything below needs a real Entra tenant, so it cannot be unit-tested. Run it once against the
tenant before trusting a deployment, and again after any change to the auth path.

Set these in the gitignored root `.dev.vars` and start the dev server (`pnpm dev-server`):

```
MICROSOFT_CLIENT_ID=<Application (client) ID>
MICROSOFT_CLIENT_SECRET=<client secret value>
MICROSOFT_TENANT_ID=<Directory (tenant) ID>
AUTH_GATEKEEPERS=microsoft
```

When a step fails, capture the `AADSTS…` code from the pop-up or the worker log — it is what
identifies the cause, and guessing without it wastes a round trip.

### Sign-in

- [ ] **Member happy path.** "Continue with Microsoft" → sign in as a tenant member on a verified
      domain → lands in the Workshop. The account's email is the address, **lowercased**; signing
      in as `First.Last@Domain.com` must produce `first.last@domain.com`, not a second account.
- [ ] **Foreign tenant rejected.** Sign in with a Microsoft account from another tenant (or a
      personal `outlook.com` account). Expected: refused with a visible reason, no account created.
- [ ] **B2B guest rejected.** Same, with a guest invited into this tenant. If the tenant has no
      guest and you will not invite one, record this item as **untested** — it is an assumption,
      not a verified control, until somebody runs it.
- [ ] **Misconfiguration is actionable.** Unset `MICROSOFT_TENANT_ID`, restart, retry sign-in. The
      pop-up must show the "Microsoft Gatekeeper Not Configured" page — a user must never be left
      watching an infinite spinner with nothing to act on. Restore the var afterwards.

### Mailbox connection

- [ ] **Connect.** Connections → Microsoft → the configurator frame shows "Outlook mailbox" →
      Connect → consent for `Mail.ReadWrite`.
- [ ] **No re-consent loop.** Reopen the configurator and connect the same mailbox again. It must
      go straight through: the granted scopes come back resource-qualified
      (`https://graph.microsoft.com/Mail.ReadWrite`) and must be recognized as already granted, not
      re-requested every time.

### Mailbox behavior (via a gadget)

- [ ] **Reads produce observations.** List and read messages; each read shows up as an observation
      in the approval feed, and nothing is written.
- [ ] **Writes queue.** Mark-read, move, and reply-draft each land in the approval queue instead of
      taking effect, and the approval description shows the real subject, sender, and draft text.
- [ ] **Immutable ids survive a move.** Approve a move of a message to another folder, then have the
      agent act on **the same message again** (mark it read, or draft a reply) and approve that too.
      Both must succeed: Outlook re-keys a moved message unless the immutable-id preference holds,
      and a broken preference shows up here as a 404 on the second action.
- [ ] **Nested folders are reachable.** Create a subfolder in Outlook (e.g. `Clients/Acme`), then
      ask the agent to list folders and move a message into it. The subfolder must appear —
      Graph's root listing omits child folders, so this exercises the traversal — and the move must
      apply.
- [ ] **A reply draft cannot inject markup.** Have the agent draft a reply whose comment contains
      `<b>bold</b> <script>alert(1)</script>`, approve it, then open the draft in Outlook. Record
      what you see: inert text (expected) or rendered HTML. If Graph interprets the comment as HTML,
      the text still cannot execute in the approval UI, but say so here and escape it at the API
      layer before this ships to anyone but you.

### Password-auth coexistence

Password accounts key on the address exactly as typed, while Microsoft sign-in lowercases it. The
same human signing up with `Alex@contoso.com` by password and then with Microsoft would end up
with two accounts. Deployments that set `DISABLE_PASSWORD_AUTH=true` (the expected staging shape)
have no exposure. If a deployment keeps password auth on alongside Microsoft, verify the address
casing your password accounts use before enabling this gatekeeper.

## Troubleshooting

### The login pop-up closes (or shows an error) and the main window spins forever

The token exchange failed. Open the pop-up's error text before it closes, or check the worker logs:

- `AADSTS7000222` — the client secret has expired. See the rotation runbook above.
- `AADSTS7000215` — the client secret value is wrong (a Secret ID was copied instead of the Value,
  or the secret was rotated without updating the deployment).
- `AADSTS65001` — admin consent was never granted for the delegated permissions (Step 2).
- `AADSTS9002325` — the app's redirect URI is registered under the "Single-page application"
  platform instead of "Web". Delete the SPA platform entry in the app registration and register the
  URI under "Web" platform only.

### "Not configured" page during authorization

One of `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID` is missing. In dev, check all three
`MICROSOFT_*` shell vars and restart the dev server.

### `AADSTS50011` / redirect URI mismatch

The app registration's **Web** redirect URI must match `${PUBLIC_BASE_URL}/gatekeeper/microsoft/oauth`
exactly — no trailing slash, and `http` (not `https`) for `localhost`.

### Sign-in is refused for a user who exists in the directory

The account is a guest, an `#EXT#` account, or its address is not on a verified domain of the
tenant. That is the identity policy above, not a misconfiguration.

### How to confirm sign-in worked

After completing the sign-in flow, verify success in two places:

1. **Dev server log**: watch `pnpm dev-server` output for a log line containing `gatekeeper login
   finished` with `"outcome":"ok"` (from the component workshop.auth).
2. **Workshop UI**: sign-in redirects to the Workshop itself (if a gadget is available) or the main
   interface. A connected Microsoft account appears under **Gatekeepers** → **Connections** or via
   the **/gatekeepers** endpoint. Note: the backend origin (localhost:8787) intentionally serves 404
   for pages; the UI is on the Vite dev client (localhost:3000).

## Build

```
pnpm --filter @gadgets/microsoft-gatekeeper build
```
