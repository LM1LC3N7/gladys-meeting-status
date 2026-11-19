# TODO — Meeting status v1

Handoff document: what is done, what was decided, what comes next. Keep it
up to date at the end of every milestone.

## Status

- [x] **Milestone 1 — core** (2026-11-15): two accounts (Microsoft Graph
      presence, Zoom presence status), OAuth2 of the user's apps with PKCE
      (Connect returning to Gladys, or a pasted address whose state may be
      wrapped by the Gladys front), tokens in `/data` renewed by a timer,
      one read per service per interval (60 s at least, guarded by /data and
      a monotonic clock), one read-only device of five features, persisted
      stops and pauses, getConfig() retried, unreadable /data read again
      every 10 s without writing over it, docs en/fr, cover. Tested against
      answers shaped like the official documentations only.
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
  signed-in service was read freshly and says 0; unknown otherwise (not
  published). A reading older than max(3 intervals, 5 minutes) is unknown.
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
