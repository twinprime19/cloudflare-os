# Microsoft 365 gatekeeper

Microsoft Entra ID (Azure AD) integration for Gadgets. It serves two purposes:

- **Sign-in:** when `microsoft` is in the deployment's `AUTH_GATEKEEPERS` allowlist, "Continue with
  Microsoft" appears on the login page. Sign-in requests only the identity scopes (`openid`,
  `profile`, `email`, `User.Read`) to read the account's **verified email**, which becomes the
  user's identity. The sign-in grant is transient (discarded right after the email is read).
- **Connections:** when a user connects Microsoft, scopes are requested for the grantable resources
  they enable (plus `offline_access`):
  - **Outlook mailbox** at `https://outlook.office.com/mail/` — `Mail.ReadWrite`, so gadgets can
    read, organize, and draft replies to mail on the user's behalf.
  - **Microsoft Teams** at `https://teams.microsoft.com/` — five read-only scopes, so gadgets can
    read the teams, channels, chats, members, and messages the user can already see. Nothing in
    this resource writes: there is no method that posts, edits, deletes, marks read, joins, or
    leaves.
  - **SharePoint List** at `https://*.sharepoint.com/*` — `Sites.ReadWrite.All`, so a gadget can
    read one list's columns and items and create new items in it. See
    [SharePoint List](#sharepoint-list).

  A connect flow that names no resources requests **every** resource's scopes, which is what the
  generic "Connect Account" button does. Every resource's permissions must therefore be consented
  before that button works — see [Step 2](#step-2-grant-the-api-permissions).

A single Entra app registration is used for both, and it is pinned to **one tenant** — the worker
needs `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` before it will answer anything.

## Identity policy

Only directory **members** on one of the tenant's **verified domains** can sign in. B2B guests
(`userType` of `Guest`, `#EXT#` accounts, or addresses on unverified domains) are rejected by
design: their addresses are owned by another tenant, so this tenant has not proven the email
belongs to the person signing in. The address is lowercased and becomes the account key, so signing
in with Microsoft resolves to the same Workshop account as any other gatekeeper that verifies the
same address.

## Profile hints: name and avatar

Sign-in also seeds the Workshop's display name and avatar from Microsoft Graph, as best-effort hints
(see [docs/oauth-signin.md](../../docs/oauth-signin.md) for the seed/backfill policy and the
never-an-identity-signal guarantee). The name comes from `/me`'s `displayName`; the photo is fetched
via a size ladder — `/me/photos/240x240/$value` → `/me/photo/$value` → `/me/photos/96x96/$value` —
accepting only JPEG/PNG bodies up to 100 KB. No additional OAuth scopes or tenant admin consent are
needed: `User.Read`, already granted for sign-in (see Step 2 in the app registration below), covers
both `/me` and `/me/photos/…`.

## Microsoft Teams: the limits of this version

What the Teams resource serves is the read side of the account's own membership: joined teams, their
channels and rosters, channel messages and replies, the user's chats with their members and
messages, and a search across them with on-demand hydration of a hit. Three limits are worth
knowing before enabling it, because none of them is a setting a deployment can change.

- **A Teams-bound gadget is single-user.** A connection covers every team, channel, and private chat
  the connected account takes part in, and no second user can be shown to have access to all of
  that — so `addObserver` throws, and the gadget is observable only by the account that connected
  it. Sharing would mean verifying each observer against their own Teams membership through their
  own delegated token and serving only the channels they belong to (private chats could never be
  shared). That is a later change, not a configuration.
- **A grant cannot be narrowed afterwards.** Reconnecting an account re-requests its full breadth,
  so there is no "keep the mailbox, drop Teams". Shedding Teams means disconnecting the Microsoft
  account altogether and connecting again naming only the mailbox.
- **Nothing is live.** Reads happen when a gadget asks for them; no arriving message wakes anything.
  Change notifications would need a public subscription endpoint renewed every few days, and the
  bulk `getAllMessages` / `/delta` feeds are application permissions — neither fits a delegated,
  on-demand gatekeeper. An agent therefore sees what it looks at, when it looks.

## SharePoint List

One connection is **one list**. A user introduces it by pasting the list's address; the gadget gets
a `SharePointListSession` with four methods — `getColumns()`, `getItems()`, `getItem()`,
`createItem()`. There is no update, no delete, no schema change, and no file or document-library
access: the intended shape is a form gadget that appends rows to a list somebody already maintains.

Every request runs on the connecting user's own delegated token, so the connection can only reach
what that user can already open in SharePoint.

### Introducing a list

Paste the address of the list page itself (the `AllItems.aspx` view is fine). The query string and
fragment are ignored — they are view state, not identity. All three site shapes resolve:

```
https://contoso.sharepoint.com/sites/HR/Lists/Requests/AllItems.aspx     → site /sites/HR
https://contoso.sharepoint.com/teams/HR/sub/Lists/Requests/AllItems.aspx → site /teams/HR/sub
https://contoso.sharepoint.com/Lists/Requests/AllItems.aspx              → the tenant root site
```

The "Copy link" button's share form (`https://contoso.sharepoint.com/:l:/r/sites/HR/Lists/Requests`)
is accepted too; its other share modes (`/:l:/s/…`, `/:l:/g/…`) name a share rather than a path and
are refused with a hint to copy the URL from the browser's address bar instead.

The list is matched on the URL name after `/Lists/` (not the display name, which anyone who can edit
the list may change). Site and list **ids** are resolved once, at introduction, so the binding keeps
working after a rename. The binding's title is the list's display name.

Three failures are worth recognizing:

- `<what is wrong>. Expected a SharePoint list URL like https://<tenant>.sharepoint.com/…` — the
  address is on a SharePoint host but does not name a list: no `/Lists/` segment, no list name after
  it, `http` instead of `https`, or a share mode that carries no path. The first sentence says which.
  No Graph call was made.
- `The Microsoft gatekeeper cannot connect this URL: …` — the address is not on a SharePoint host at
  all, so it is not one this gatekeeper connects. No Graph call was made.
- `Could not open that SharePoint list: … Check the URL and that your account can open it.` — the
  address parsed, but Graph would not resolve the site or the list for this account. The middle part
  is Microsoft's own reason.

### Columns a gadget can read and write

`getColumns()` reports each column's **internal name** (the key `createItem()`, `select`, and
`where` all use), display name, type, `required`, `readOnly`, a `choices` list for choice columns,
`multiline` for a text column that takes more than one line, and `multiple` for a person or lookup
column that holds more than one value. Hidden, system, and computed columns are not reported at all
— except `Title`, which some lists mark read-only in a view while still requiring it on create.

| Type | Read | Write |
| --- | --- | --- |
| `text` (incl. multiline), `number`, `boolean`, `dateTime`, `choice` | yes | yes |
| `person`, `lookup` | yes | **no** — a numeric directory/list id this cannot look up |
| `unsupported` (currency, calculated, hyperlink, managed metadata, geolocation, …) | yes | **no** |

`createItem()` validates against the list's own schema **before** anything is submitted, so a bad
call fails immediately instead of sitting in the approval queue. It refuses: an argument that is not
an object; an unknown column name; a missing (or explicitly `null`) required column; a value of the
wrong type; a `choice` outside the column's options; and any `person`, `lookup`, or `unsupported`
column. Dates are normalized to an ISO timestamp, so neither the caller's formatting nor its time
zone reaches SharePoint.

Reads: `getItems()` returns a cursor — 25 items per page by default, 200 at most, at most 40 pages
per cursor. Filters are structured (`{column, op, value}`, at most five clauses, combined with
`and`); there is no way to pass a raw OData string. Every read is recorded as an observation in the
workspace's activity, and nothing a read does can change the list.

### The "Create list items" auto-approval rule

By default each `createItem()` lands in the approval queue with a fenced key/value description of
the values, and a build collaborator approves it before the row is written.

Turning on the **Create list items** rule makes those submissions apply without review. Read this
before enabling it:

> **The rule is keyed to the introduced list in the workspace, not to a gadget.** Once it is on, it
> covers **every gadget in that workspace bound to that list — current and future**. A second gadget
> bound to the same list writes rows with no approval step, and so does the next one somebody builds
> there. If that is not wanted, keep the rule off, or keep the list in a workspace whose build
> collaborators you trust with it.

Enabling it also clears the backlog: creates already waiting for this list are applied the moment
the rule goes on. Disabling it only affects later submissions — whatever is already queued still
needs a human.

### What `createItem()` gives back

With the rule on and nothing else in flight, `createItem()` waits up to **5 s** for its own action
to be applied and then returns the real item id — `{id: "218", provisional: false}`. Otherwise it
returns a provisional id — `{id: "pending-7", provisional: true}` — which is the gatekeeper's own
bookkeeping for this list, not a SharePoint id. It is held in the binding's storage, so it keeps
resolving from any later session; store it rather than discarding it when a session ends. The drain
applies one action at a time per list, so three submissions a second apart will usually see the
first answer with a real id and the rest with provisional ones.

A submission decided **against** inside those 5 s never comes back as a provisional id:
`createItem()` itself throws the same `SharePointSubmissionError` — rejected, failed, or throttled —
that `getItem()` would have thrown later. Handle those three at the call that submits, not only at
the lookup.

Pass a provisional id to `getItem()` later to find out what happened. Six outcomes, five of which
are errors:

- **applied** — the real item is returned, exactly as if you had passed its SharePoint id.
- `Submission <n> was rejected by an approver.` — a `SharePointSubmissionError` with
  `outcome: "rejected"`.
- `Submission <n> failed: <what SharePoint said>` — `outcome: "failed"`, `throttled: false`.
- `Submission <n> failed: SharePoint throttled the write. It is still queued, so an approver can
  retry it from the workspace's activity list.` — `outcome: "failed"`, `throttled: true`.
- `Submission <n> is queued behind an earlier failed submission; a build collaborator must clear it
  from the workspace's activity list.` — see below.
- `Submission <n> has not been applied yet.` — still waiting for approval, nothing is wrong.

A binding holds at most **100** pending actions; the 101st submission is refused with "Too many
pending SharePoint list actions. Resolve existing actions before adding more."

### When a submission gets stuck

A create whose Graph write failed **stays pending on purpose**: the drain stops at it rather than
skipping ahead, so a human decides whether to retry or drop it. Until then, every later submission
for that list reports "queued behind an earlier failed submission" — including other people's.

Clearing it: a **build** collaborator opens the workspace's **Activity** panel and approves the
action again (a retry, for a throttle) or rejects it. Collaborators with the **use** role get the
minimal UI — top bar and gadget only — so they cannot see the stuck action and cannot clear it. A
workspace shared for use therefore needs at least one build collaborator who watches Activity.

### Sharing a gadget

A collaborator is admitted as an observer only if **their own** Microsoft account can open the list;
the check runs on their token, not the owner's. A collaborator without site access gets "You do not
have access to this SharePoint list." If Microsoft cannot be reached to answer the question, they
get "Could not verify SharePoint access right now." and can be retried later.

### Known limits

- **Duplicate window after a crash.** If the worker dies between SharePoint accepting the row and
  the gatekeeper recording that it did, the row exists while the action still looks pending —
  approving it again creates a second row. This is accepted deliberately: the alternative was
  requiring a marker column in every list this ever writes to. Before re-approving an action that
  failed without a clear reason, look at the list.
- **Outcome records are never swept.** One small record per submitted action stays in the binding's
  storage for the binding's lifetime. Harmless at form volumes; a retention policy is a future
  decision, not a current control.
- **A column added in SharePoint is not picked up while the connection is warm.** The list's schema
  is cached in memory per Durable Object instance, and the only thing that invalidates it is a Graph
  4xx at apply time. Since `createItem()` validates against the cached schema *before* submitting, a
  newly added column never reaches that path — it is refused with `Unknown column "X". Use internal
  column names from getColumns().`, and `getColumns()` keeps returning the old schema, until the
  instance goes idle. After adding a column, leave the workspace alone for a few minutes before
  writing to it.
- **Reads have their own throttling ceiling.** A read waits out a SharePoint `Retry-After` for at
  most 10 s (writes get 60 s, since an approval queue rather than a caller is waiting). Past it
  `getColumns()`, `getItems()` and `getItem()` fail with `SharePoint is throttling this connection
  and asked to be left alone for N seconds. Try again after that.` Nothing is lost — retry after the
  stated wait.
- **Multi-select choice columns look single-valued.** Graph's column definition exposes no
  multi-select flag, so such a column is described as an ordinary `choice` and only accepts one
  string. A multi-value write is expected to be refused by SharePoint — untested, because no list
  with such a column has been tried yet. Avoid those columns until one has.
- **Large lists.** SharePoint's 5000-item list view threshold applies to `getItems()`: past it,
  filtering or sorting on a column that has no index fails, and Microsoft's own sentence comes back
  as `SharePoint rejected the request: Field 'X' cannot be referenced in filter or orderby as it is
  not indexed`. The fix is on the SharePoint side — add an index for that column in list settings.
- **Not yet verified against a real tenant** (check these on the first live list; they are
  assumptions the code makes): the percent-encoded `$expand=fields($select=…)` form, quoted
  `dateTime` literals in filters, the `$top` ceiling of 200, and root-site resolution for a
  tenant-root list.

### Rollback

Withdrawing SharePoint is the same procedure every resource here uses — withdraw it in code, never
touch the migration: see
[Durable Object migrations, and how to back out a resource](#durable-object-migrations-and-how-to-back-out-a-resource).
The class to keep exported is `SharePointListGatekeeperImpl` and the tag to keep is
`v1-sharepoint-list`.

De-consenting `Sites.ReadWrite.All` is a separate action in Entra (API permissions on the app
registration) and affects nothing in this repo. Note that it also breaks the generic "Connect
Account" flow, which asks for every resource's scopes.

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

| Permission | Admin consent | Why |
| --- | --- | --- |
| `openid`, `profile`, `email` | no | sign-in and the id token claims the email is read from |
| `offline_access` | no | refresh tokens, so a connected mailbox survives past the first hour |
| `User.Read` | no | the signed-in user's profile, and the tenant's verified domain list |
| `Mail.ReadWrite` | no | read messages and folders, flip read state, move messages, draft replies |
| `Team.ReadBasic.All` | no | list the teams the user has joined, and read a team's name |
| `Channel.ReadBasic.All` | no | list and read the channels of those teams |
| `TeamMember.Read.All` | **yes** | read a team's roster |
| `ChannelMessage.Read.All` | **yes** | read channel messages and their replies |
| `Chat.Read` | no* | read the user's chats, their members, and their messages |
| `Sites.ReadWrite.All` | **yes** | SharePoint List: read a list's columns and items, and create items |

Then click **Grant admin consent for &lt;tenant&gt;** and confirm every permission reads
**Granted**. Many tenants block user consent; without admin consent those users hit `AADSTS65001`
("The user or administrator has not consented") on their first sign-in and can never get past it.

`Sites.ReadWrite.All` is broad by necessity: Graph's per-site permission (`Sites.Selected`) is
application-only, so a *delegated* connection cannot be scoped to one site. Every request still runs
as the signed-in user, so the connection reaches nothing that user could not already open — and a
gadget only ever sees the single list its binding names.

**Consent all of a resource's permissions or none.** A resource counts as granted only when the
grant covers *every* one of its scopes, so a resource left unconsented is unavailable rather than
partly working — and the symptom is a connect flow that keeps offering a reconnect link. The worker
logs the shortfall at `warn` with `event: "microsoft.consent.resource.ungranted"`, naming the
resource and the missing scopes; check the logs before re-running the consent. A grant records what
it covered at the moment it was issued, so **an account that connected before the permission was
added stays without that resource until it connects Microsoft again** — fixing the consent alone
changes nothing for accounts already connected.

**Order of operations.** Because a connect flow that names no resources asks for every resource's
scopes, the permissions above must be consented in the tenant **before** a build carrying a new
resource reaches users. Otherwise the generic "Connect Account" button fails with `AADSTS65001` for
everyone, including users who only wanted mail. Adding a permission to a registration that is
already deployed is the same place — **API permissions** > **Add a permission** > **Microsoft
Graph** > **Delegated permissions** — followed by **Grant admin consent** again; re-granting does
not disturb the grants already in use, and no redeploy is needed for the consent itself.

**\*The "admin consent" column is Microsoft's static classification of the permission, not a
prediction about this tenant.** Under "Secure by Default" (message center post `MC1163922`) it is
the tenant's **consent policy** that decides what a user may agree to, and the default policy
withholds permissions that read other people's data. A "no" row can therefore still stop at a
prompt only an administrator can complete — `Chat.Read` most of all, since chat messages are by
definition other people's. Granting admin consent for the whole table once makes the distinction
moot, which is why the step above is not optional. **Record this tenant's actual posture here**
when the permissions are added: one line saying whether a user could consent unaided, or the
administrator had to.

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

### SharePoint list (via a gadget)

Run this against a real list you are allowed to write to — every item it creates is a real row that
the site's members can see.

- [ ] **Consent is in place.** On the app registration this deployment uses, `Sites.ReadWrite.All`
      (delegated) reads **Granted**. Each environment with its own registration needs its own grant,
      before the build that requests the scope is deployed there.
- [ ] **Connect.** Connections → Microsoft → the picker offers "SharePoint List" → the frame asks
      for a list URL, and **Connect** stays disabled until the address is plausible. Paste a
      `/sites/…/Lists/…` URL → a binding appears, titled with the list's display name.
- [ ] **Other site shapes resolve.** Repeat with a `/teams/…` list and with a list on the tenant
      root site (no `/sites/` or `/teams/` segment). Both must bind.
- [ ] **A bad address fails cleanly.** Paste a non-list URL → a readable error, no binding, and
      nothing left half-connected.
- [ ] **Schema is visible.** Ask the agent to describe the binding and call `getColumns()` → types,
      required flags, and a choice column's options are all shown; hidden system columns are not.
- [ ] **Auto-apply returns a real id.** Enable the **Create list items** rule for this list, then
      submit once from a gadget → the row appears in SharePoint within seconds and the gadget shows
      a real item id, not `pending-<n>`.
- [ ] **Contended submissions still land.** Submit three times within a few seconds → all three rows
      appear; the later ones may answer with provisional ids, and `getItem()` on those must resolve
      to the real rows.
- [ ] **The rule's blast radius is real.** Bind the same list into a second gadget in the same
      workspace → it writes without any approval. This is the documented behaviour; confirm you are
      happy with it before enabling the rule anywhere that matters.
- [ ] **Without the rule, writes queue.** Disable the rule and submit → the action waits in
      **Activity** with a readable, fenced description of the values; approve → the row is created.
- [ ] **Validation happens before the queue.** Submit with a required field missing → refused
      immediately, naming the column, with nothing left in the approval queue.
- [ ] **Types round-trip.** Write a choice, a number, a boolean, and a date → each is stored with
      the right type in SharePoint (a date lands as the moment you meant, not a day out).
- [ ] **Sharing respects SharePoint's own access.** Share the gadget with a colleague who has site
      access → they can use the form. Share it with someone who does not → they are refused with
      "You do not have access to this SharePoint list."
- [ ] **Filtering works on a real list.** `getItems()` with a `where` clause on an indexed column
      and `top` of 200 → pages correctly. Record the largest `top` Graph actually honoured; on a
      list past the 5000-item threshold, filtering an unindexed column is expected to fail with
      Microsoft's "not indexed" sentence.

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
- `AADSTS65001` — admin consent was never granted for the delegated permissions (Step 2). It is
  also what a generic "Connect Account" looks like while a newly added resource is unconsented:
  that flow names no resource, so it asks for every resource's scopes and fails for everyone,
  including a user who only wanted mail.
- `AADSTS9002325` — the app's redirect URI is registered under the "Single-page application"
  platform instead of "Web". Delete the SPA platform entry in the app registration and register the
  URI under "Web" platform only.

### "Not configured" page during authorization

One of `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID` is missing. In dev, check all three
`MICROSOFT_*` shell vars and restart the dev server.

### `AADSTS50011` / redirect URI mismatch

The app registration's **Web** redirect URI must match `${PUBLIC_BASE_URL}/gatekeeper/microsoft/oauth`
exactly — no trailing slash, and `http` (not `https`) for `localhost`.

### Every Graph call fails while the account still looks connected

Conditional Access — or continuous access evaluation — rejected the token's claims, and Graph
answers `401` with `WWW-Authenticate: Bearer … error="insufficient_claims"`. Minting again cannot
help, since the refresh token produces a token with the same claims, so the gatekeeper does not
retry: it reports the credentials as dead, which is what raises the reconnect prompt. The user
reconnects and satisfies whatever the policy now asks for (MFA, a compliant device). This is one
shared request path, so Teams behaves exactly as mail does here.

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

## Durable Object migrations, and how to back out a resource

Each resource type has its own Durable Object class, and every class is registered by a `migrations`
tag in `wrangler.jsonc` — `v0` for `UserAccount` and `OutlookMailGatekeeperImpl`, `v1` for
`TeamsGatekeeperImpl`, `v1-sharepoint-list` for `SharePointListGatekeeperImpl`. Tags are appended,
never edited: rewriting an applied tag makes the next deploy disagree with the migration history the
account already holds.

**A deployment that has applied a tag cannot be rolled back past it.** `wrangler rollback` refuses
to cross a Durable Object migration, and redeploying a pin from before that tag fails the same way,
because the account's history records a class the older bundle does not declare. Recovery is
therefore forward, not backward.

**To withdraw a resource from users without redeploying an older build**, remove three things and
ship that:

1. its entry in `RESOURCE_SCOPES` (`src/microsoft.ts`) — it stops being offered, stops being
   consented for, and disappears from the connect UI;
2. its branch in `getGatekeeperClassFor` — no URL routes to it;
3. its branch in `startResourceConfigurator` — no frame is served for it.

**Keep the Durable Object class exported and keep its migration tag in place.** They are what the
migration history refers to; dropping either turns the next deploy into the same mismatch a rollback
would have caused. The resource is then invisible to users, existing objects are inert, and the
other resources are untouched.
