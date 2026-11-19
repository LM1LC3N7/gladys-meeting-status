# Changelog

All notable changes to this integration are documented in this file. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
uses [semantic versioning](https://semver.org/).

Describe each change under `## [Unreleased]` as you make it. The Release
workflow moves that section under the version it ships, and the section becomes
the notes of the version's GitHub Release.

## [Unreleased]

### Added

- "In a meeting" status of Microsoft Teams (Microsoft Graph presence, work or
  school accounts) and Zoom (presence status), read once a minute at most:
  one read-only device with "In a meeting (Teams or Zoom)", one value per
  service and the presence texts.
- OAuth2 with the user's own apps (Microsoft Entra ID, Zoom Marketplace),
  PKCE, Connect returning to Gladys or a pasted address, tokens renewed before
  their end and kept in `/data` only.
- Option to count busy and do not disturb as "in a meeting".
- Pauses kept across restarts (429 `Retry-After`, 5xx, 403, daily budget),
  refusals kept until a new Connect, an unknown presence never published.
