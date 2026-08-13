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
5. **Redirect URI**: platform **Web**, value `${PUBLIC_BASE_URL}/gatekeeper/microsoft/oauth` — for
   local dev that is `http://localhost:8787/gatekeeper/microsoft/oauth`.
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

1. Pick an expiry. Entra's maximum is **24 months**; the console's default is commonly **180 days**.
   Longer means fewer rotations; shorter limits how long a leaked secret is useful. Pick
   deliberately — this is not a value to accept by reflex.
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
`run-dev-server.js` maps them into the worker:

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

_Filled in by the manual verification phase._

## Troubleshooting

### The login pop-up closes (or shows an error) and the main window spins forever

The token exchange failed. Open the pop-up's error text before it closes, or check the worker logs:

- `AADSTS7000222` — the client secret has expired. See the rotation runbook above.
- `AADSTS7000215` — the client secret value is wrong (a Secret ID was copied instead of the Value,
  or the secret was rotated without updating the deployment).
- `AADSTS65001` — admin consent was never granted for the delegated permissions (Step 2).

### "Not configured" page during authorization

One of `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID` is missing. In dev, check all three
`MICROSOFT_*` shell vars and restart the dev server.

### `AADSTS50011` / redirect URI mismatch

The app registration's **Web** redirect URI must match `${PUBLIC_BASE_URL}/gatekeeper/microsoft/oauth`
exactly — no trailing slash, and `http` (not `https`) for `localhost`.

### Sign-in is refused for a user who exists in the directory

The account is a guest, an `#EXT#` account, or its address is not on a verified domain of the
tenant. That is the identity policy above, not a misconfiguration.

## Build

```
pnpm --filter @gadgets/microsoft-gatekeeper build
```
