# CLAUDE.md

Project rules for coding assistants (Claude Code reads this file) and for human
contributors.

## What this is

A Gladys Assistant **external integration** of type `device`: a Node.js
service (Node 22+, ESM, no build step) in its own sandboxed Docker container,
talking to Gladys through `@gladysassistant/integration-sdk` (v0.16), its only
runtime dependency (HTTPS goes through `node:https`). It reads the presence of
the user in Microsoft Teams (Microsoft Graph `GET /me/presence`, work or school
accounts) and Zoom (`GET /users/me/presence_status`), read only, after an
OAuth2 authorization (PKCE) of apps the user registers. Never tested with a
real account. Read **[TODO.md](./TODO.md)** first: status and decisions.

## Commands

```bash
npm install
npm test                                  # node --test, the built-in runner
node --test test/app.test.js              # one file
node --test --test-name-pattern "plug"    # tests whose name matches
npm run lint                              # ESLint
npm run format:check                      # Prettier, as the CI runs it
npm run format                            # Prettier, fix in place
npx github:GladysAssistant/integration-store .   # store admission checks
```

The CI (`.github/workflows/ci.yml`) runs `format:check`, `lint` and `test` on
Node 22 and 24, and builds the Docker image: run the three commands before
pushing.

## Architecture

```
index.js            SDK bootstrap only
src/handlers.js     every SDK handler -> the app (shared with the tests)
src/app.js          MeetingApp: /data, getConfig retry, the two accounts, device, states,
                    badge, status, Connect buttons, pasted address, Test button
src/account.js      Account: one service: session, renewal timer, reads, pauses, stops
src/oauth.js        authorization addresses (PKCE), token endpoints, pending sign-ins
src/presence.js     the presence read and what it means (in a meeting or not, unknown)
src/http.js         the only network access (node:https, bounded), errors by status
src/devices.js      the device and its states (pure)
src/store.js        /data/meeting-status-state.json (0600): stops, pauses, times (never a token)
src/config.js       config_schema normalized, client key (hash of service, id, tenant)
src/constants.js    endpoints, keys, limits, timings
src/i18n.js         runtime texts, device and feature names (en/fr)
scripts/cover.mjs   cover.png (800x534) in plain Node
scripts/mutants.mjs guards broken one by one, in a private copy (dev VM)
test/helpers/       fake Microsoft and Zoom behind the transport; Gladys is the SDK's fake
test/kit/           clocks and HTTPS test servers copied from gladys-testkit
```

## Rules

- **External ids are forever.** Build them with
  `gladys.externalIds(type, platformId)`, `platformId` being the unique id the
  platform gives the device (serial, cloud id, MAC), never a label. Gladys keys
  devices, their history, scenes and dashboards on them: changing one creates a
  new device and orphans the old one. The same goes for every key users store:
  manifest `config_schema`, `actions`, widget and scene keys, feature keys.
- **The container rootfs is read-only.** Write only under `/data`, the one
  writable volume. Never log tokens, passwords or API keys.
- **Versions belong to the Release workflow** (Actions → Release). Never edit
  `version` in `package.json` or in the manifest, nor the `docker_image` tag,
  by hand — except on the README's hand-pushed tag path, which bumps all three
  together before tagging. `test/manifest.test.js` checks they agree.
- **Changelog.** Every user-visible change adds a line under `## [Unreleased]`
  in `CHANGELOG.md` (`### Added`, `Changed`, `Fixed`, `Removed`, `Security`).
  Never write a version heading: the Release workflow moves the section, and it
  becomes the notes of the GitHub Release.
- **Manifest.** Texts are multi-language objects `{ "en": ..., "fr": ... }`:
  `label`, `description` and `placeholder` alike, never a plain string. The
  catalog `description` holds 10 to 100 characters per language. A field the
  older cores do not know needs a higher `gladys_version` minimum
  (`categories`: 4.86.0; `widgets`, `scene_triggers`, `scene_actions`,
  `type: "provider"`: 5.1.0). `test/manifest.test.js` ties the manifest to the
  code (defaults, action, scene and widget handlers): change both sides
  together.
- **User-facing text** is bilingual (English and French): action messages,
  manifest texts, and `docs/en.md` / `docs/fr.md`, kept in sync.
- **Style.** Prettier formats, ESLint catches mistakes. Comments explain why,
  in English. Tests never touch the network.
- **Secrets**: never a token, a code nor a client secret in a log, a status, an
  error message or a state; the tokens live in the configuration of Gladys only (`setConfig`, key
  `<service>_session` outside the schema), never in `/data`. A secret field is
  described as "kept by Gladys in its database, never shown again" (Gladys does
  not protect it further: never claim more).
- **Renewal tokens serve once**: a renewal leaves only after Gladys accepted a
  witness write (`saveSession({ witness: true })`), a failed save is retried by
  `scheduleSessionSave` (bounded, cleared by `stop()`), writes go one at a time
  (`sessionChain`), and `shutdown()` waits (8 s at most) for a renewal in flight
  and its save; `closing` stops new renewals and reads, `closed` everything.
- **Never invent a value**: an unknown presence is not published, never as 0.
- **One read per service per interval** (a minute at least), whatever starts it.
