# TODO — Meeting status v1

Handoff document: what is done, what was decided, what comes next. Keep it
up to date at the end of every milestone.

## Status

- [x] **Milestone 1 — core** (2026-11-15): two accounts (Microsoft Graph
      presence, Zoom presence status), OAuth2 of the user's apps with PKCE
      (Connect returning to Gladys, or a pasted address whose state may be
      wrapped by the Gladys front), tokens renewed by a timer,
      one read per service per interval (60 s at least, guarded by /data and
      a monotonic clock), one read-only device of five features, persisted
      stops and pauses, getConfig() retried, unreadable /data read again
      every 10 s without writing over it, docs en/fr, cover. Tested against
      answers shaped like the official documentations only.
- [x] **Fixes of review 1** (2026-11-19): a stopped service stays in "Teams
      or Zoom" (unknown, never 0 by omission); shutdown during an exchange or
      a renewal leaves no timer nor request; tokens moved to Gladys; freshness
      3 minutes at least; tests and mutants for the gaps found; CI jobs with
      `timeout-minutes`.
- [x] **Renewed tokens never lost** (2026-11-19, model of gladys-oura 0.1.3):
      renewal only after a witness write Gladys accepted, failed saves retried
      by a short bounded timer and at each reconnection, shutdown waits up to
      8 s for a renewal in flight and its save; tests and mutants.
- [ ] **Milestone 2 — real accounts**: a Microsoft 365 work account and a Zoom
      account; check the exact Zoom status spellings returned, the Microsoft
      consent errors, the redirect through my.gladysassistant.com.
- [ ] **Later, if asked**: Zoom webhooks through Gladys Plus (presence pushed
      instead of read), Microsoft Graph change notifications, Webex.

## Decisions

- **Version 0.1.0 set by hand** in `package.json`, `package-lock.json`, the
  manifest and `docker_image` (`ghcr.io/lm1lc3n7/gladys-meeting-status:0.1.0`).
- **No reference code**: written from the official documentations (Microsoft
  Graph `presence-get`, Microsoft identity platform v2.0, Zoom Users API
  "Get a user presence status", Zoom OAuth). Nothing to port: no
  THIRD_PARTY code.
- **Model**: the account logic is the one of gladys-oura (approved): stops
  kept in /data, pauses, renewal timer with a one-minute floor, a single
  flight per account, the Test button as the only way to retry a refused
  access.
- **One device, fixed id** `meeting:status`: a meeting light keeps its device
  when the user changes app or account. Features `unknown`/`binary` (never a
  switch: nothing to command) and `text`/`text`.
- **"In a meeting (Teams or Zoom)"**: 1 when one says 1; 0 only when every
  authorized service (stopped ones included) was read freshly and says 0;
  unknown otherwise (not published). A reading older than max(3 intervals,
  3 minutes) is unknown.
- **Unknown values**: a presence the documentation does not list is unknown,
  never 0. The documented spellings of Zoom are all known (the doc lists two
  sets: `In_A_Zoom_Meeting`/`In_A_Meeting`, `On_A_Call`/`In_A_Call`…).
- **Refusal**: only a 400/401 of a token endpoint carrying an OAuth `error`
  code; a 400/401/403 without one (proxy, firewall) is an ordinary error.
  `temporarily_unavailable` and `server_error` are transient.
- **403 on the presence** (and Zoom code 5411): a pause of an hour, lifted by
  a new Connect; never a stop (the tokens work).
- **Tenant**: `organizations` by default; `common`/`consumers` refused
  (personal accounts have no presence). Part of the session key: another
  tenant asks for a new Connect.
- **Tokens in Gladys, never in /data** (review 1, rule of the brief): saved by
  `setConfig({ <service>_session })` (keys outside the schema, never sent to
  the front), before the new access token is used. A renewal is sent only
  while Gladys is connected and just accepted a witness write of the current
  session (else retried after 1 min, never a loop); a failed save is retried
  by a timer (10 s, 10 s, 10 s, 30 s, then every minute), at each
  reconnection and before each read; one write at a time (each sends the
  session of the moment it leaves). Shutdown (wired in src/handlers.js) sets
  `closing` (no renewal, read nor code exchange leaves), waits up to 8 s for
  a renewal in flight, one last write of a token still pending and the
  writes, then `closed`. Left: a crash or kill in the seconds of a renewal,
  or a stop while Gladys still refuses the new token.
- **Known, accepted (review r1 of the token fix)**: the session write of a
  code exchange goes through the write queue, so behind a witness write or a
  retry in flight (SDK timeout 15 s) the OAuth answer to Gladys may pass its
  5 s ack while the session is saved; unlikely, nothing lost. /data keeps stops, pauses and times only. Tokens of an older
  /data are moved to Gladys at start, then removed from /data (kept there if
  Gladys refused them).
- **Daily budget**: 2,000 requests per service per UTC day (normal pace about
  1,500).
- **States**: `publishChangedStates`, no heartbeat (a heartbeat would date an
  old value as new). In-meeting keeps its history; presence texts do not.
- **No `keywords`** in the manifest; category `services`; no `countries`
  (worldwide).
- **Teams personal / Teams free / Google Meet**: not covered, said in the docs.

## To check on a real account

- Microsoft: the presence of a user in a call from the mobile app; the error
  code of a personal account at the authorize page; refresh token rotation.
- Zoom: whether the granular scope `user:read:presence_status` is enough on a
  general user-managed app; the status of a user in a meeting started from the
  web client; the refresh token rotation (the old one invalid at once).
- [ ] r2 mineurs : effacer les jetons de la config Gladys quand les clés sont vidées ou changées (configUpdated) ; ne plus écrire legacySession: null dans /data ; tests pending-never-cleared, migrate-overwrites-newer, signedIn-block-any-client, borne fresh-min-5.

## Minors from review meeting-status-tokens-r2 (2026-11-19)

- A Connect refused during the stop consumes its OAuth state: the user starts again (the message says so).
- Two refusal log lines can come close together when the last save at stop follows a failure.
