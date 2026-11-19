# Changelog

All notable changes to this integration are documented in this file. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
uses [semantic versioning](https://semver.org/).

Describe each change under `## [Unreleased]` as you make it. The Release
workflow moves that section under the version it ships, and the section becomes
the notes of the version's GitHub Release.

## [Unreleased]

### Fixed

- A renewed token is never lost when Gladys cannot keep it: a renewal is sent
  only while Gladys is connected and has just accepted a write of the current
  tokens (otherwise tried again a minute later), and a new token Gladys
  refused is saved again within seconds and at each reconnection to Gladys,
  not only at the next read. Stopping the container waits up to 8 s for a
  renewal in flight and its save, and no renewal starts during the stop.

## [0.1.1] - 2026-11-19

### Added

- "In a meeting" status of Microsoft Teams (Microsoft Graph presence, work or
  school accounts) and Zoom (presence status), read once a minute at most:
  one read-only device with "In a meeting (Teams or Zoom)", one value per
  service and the presence texts.
- OAuth2 with the user's own apps (Microsoft Entra ID, Zoom Marketplace),
  PKCE, Connect returning to Gladys or a pasted address, tokens renewed before
  their end and kept in `/data` only.
- The tokens are kept by Gladys in its database, never in `/data`.
- Option to count busy and do not disturb as "in a meeting".
- Pauses kept across restarts (429 `Retry-After`, 5xx, 403, daily budget),
  refusals kept until a new Connect, an unknown presence never published.

[Unreleased]: https://github.com/LM1LC3N7/gladys-meeting-status/compare/v0.1.1...HEAD
[0.1.1]: https://github.com/LM1LC3N7/gladys-meeting-status/releases/tag/v0.1.1
