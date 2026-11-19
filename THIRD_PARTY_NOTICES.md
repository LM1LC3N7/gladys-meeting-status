# Third-party notices

This integration contains no code taken from another project. It calls the public APIs of
Microsoft (Microsoft identity platform, Microsoft Graph `GET /me/presence`) and Zoom (OAuth,
`GET /users/me/presence_status`), written from their official documentations only.

Its only runtime dependency is `@gladysassistant/integration-sdk` (Apache-2.0), installed by npm
with its own license.

The test clock and test servers under `test/kit/` are copied from gladys-testkit (Apache-2.0, same
author as this integration). They are not in the Docker image.

Microsoft, Microsoft Teams, Microsoft Graph and Entra are trademarks of Microsoft Corporation;
Zoom is a trademark of Zoom Communications, Inc. They are used here only to name the services the
integration reads.
