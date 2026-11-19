# gladys-meeting-status

[Gladys Assistant](https://gladysassistant.com) external integration: your **"in a meeting"**
status in **Microsoft Teams** (Microsoft Graph presence) and **Zoom**, read about once a minute, to
switch on a light outside your office from a scene. Read only.

- One device, **Meeting status**: "In a meeting (Teams or Zoom)", one in-meeting value per service
  and the presence text of each service.
- OAuth2 with **your own app** (Microsoft Entra ID, Zoom Marketplace), PKCE, tokens renewed before
  their end and kept by Gladys in its database (never in `/data`).
- Microsoft: **work or school accounts only** (Microsoft Graph has no presence for personal
  accounts). Google Meet has no presence API: not covered.
- One read per service per minute at most; 429 `Retry-After`, growing pauses on 5xx, refusals kept
  until you reconnect.

User guide: [English](docs/en.md) · [Français](docs/fr.md).

Unofficial integration, not affiliated with Microsoft nor Zoom. Never tested with a real account.

## Development

```bash
npm install
npm test            # node --test, against the SDK's fake Gladys and fake services
npm run lint
npm run format:check
npm run audit
npm run mutants     # the guards broken one by one: the suite must fail (dev VM only)
```

Architecture and rules: [CLAUDE.md](CLAUDE.md). Status and decisions: [TODO.md](TODO.md).

## Release

Versions are set by the **Release** workflow (Actions → Release): never edit `version` in
`package.json`, the manifest or the `docker_image` tag by hand. The workflow moves the
`## [Unreleased]` section of [CHANGELOG.md](CHANGELOG.md), tags, builds the multi-arch image
(`linux/amd64`, `linux/arm64`) to GHCR and publishes the GitHub Release.

## License

[Apache-2.0](LICENSE). No third-party code is included (see
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)).
