# AI models

AI models belong to the deployment, not to individual users. An admin curates one catalog on
**Providers** (`/providers`, admin-only), and every user chats with what is in it.

## The shared catalog

`AdminApi.addSharedModel()` stores a model's profile plus its provider configuration on the
`AdminSettings` durable object. Each user's own object reads that catalog when it lists models, so a
model an admin adds is immediately available to everyone, in the chat picker and in workspace AI-model
bindings alike.

A model's configuration never leaves the server. `listModels()` and the admin panel both return
profiles only — `{type, id, name}` — so a provider API token cannot reach a client, admin included.
Editing a model means supplying its token again.

In AI Gateway mode the gateway's built-in models are listed first, then the shared ones. A shared
model may not reuse a built-in's id: the built-in is what a chat turn would resolve, so the shared
entry would name a model nobody could reach, and `addSharedModel()` rejects it.

## The default model is forced

`Providers` → **Set as default** names the model every new chat starts on. It is not a suggestion:
the backend moves it to the front of every user's model list, and the frontend starts each new chat
on the first model listed.

- Nothing is remembered client-side. There is no stored model preference, in any browser.
- A user may switch models inside a chat from the admin's list; the switch applies to that chat only.
- An existing chat keeps the model it was started with.
- Externally delivered messages follow the same rule: the chat's own model, otherwise the default.
- Deleting the default model clears the setting. A default naming a model that is gone is ignored,
  so the list falls back to natural order rather than breaking.

Leaving the default unset is fine — the catalog is then listed in its natural order and a new chat
starts on whatever is first.

## Per-user models are gone

This is a breaking change. Models a user added for themselves before this change are no longer listed
and can no longer be chatted with. An old chat pinned to one fails its next turn with "No such
model"; the chat's model picker still offers the shared catalog, so the user picks again from there.
`AuthenticatedApi.addModel()` now throws for anyone who is not an admin, and the per-user
model store is no longer read by anything. Nothing migrates it: an admin re-adds the models the
deployment should offer.

## Quick model

The "quick model" (fast, cheap tasks such as generating chat titles) is not an admin setting.

- **AI Gateway mode**: hardcoded to a small fast model. Nothing to configure, titles work.
- **Otherwise**: there is no quick model, because the only store that held one was per-user.
  Automatic chat titles are skipped; everything else is unaffected.

## Deleting a model does not revoke its token

`deleteSharedModel()` removes the catalog entry. A workspace that already bound the model holds its
own snapshot of the configuration, API token included, taken when the binding was made — deletion does
not reach into it. **To actually revoke access, rotate the key at the provider.**
