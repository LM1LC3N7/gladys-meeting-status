# Meeting status (Microsoft Teams and Zoom)

Gladys reads, about once a minute, whether you are in a meeting or a call, in **Microsoft Teams**
(through Microsoft Graph) and in **Zoom**, so a scene can switch on a light outside your office.
Read only: nothing is changed on your accounts.

Unofficial integration, not affiliated with Microsoft nor Zoom. Never tested with a real device
nor a real account: tested against answers shaped like the official documentations.

## What you get

One Gladys device, **Meeting status** (Discovery tab, once a service is authorized), read only:

| Feature                      | Value                                                                         |
| ---------------------------- | ----------------------------------------------------------------------------- |
| In a meeting (Teams or Zoom) | 1 when one of the services says so; 0 when every authorized service says no   |
| In a Teams meeting           | 1 / 0                                                                         |
| Teams presence               | the codes of Microsoft Graph, `availability / activity` (`Busy / InAMeeting`) |
| In a Zoom meeting            | 1 / 0                                                                         |
| Zoom presence                | the code of Zoom (`In_A_Zoom_Meeting`, `Do_Not_Disturb`…)                     |

"In a meeting": a meeting, a call or a presentation (Teams `InAMeeting`, `InACall`,
`InAConferenceCall`, `Presenting`; Zoom `In_A_Zoom_Meeting`, `On_A_Call`, `Presenting`…). With the
option **Also busy and do not disturb**, Busy, Do not disturb and the Zoom calendar events count too.

A value is never invented: a presence the documentation does not list (`PresenceUnknown`, a new
value), or a service that could not be read for more than 3 intervals (3 minutes at least), is
**unknown** and nothing is published. Gladys has no "unknown" value: **the last value stays shown**;
read the integration status (or the "Teams presence" / "Zoom presence" texts and their date)
before trusting the light. "In a meeting (Teams or Zoom)" goes to 0 only when every authorized
service was read lately and says no: a service whose session ended (renewal refused) stays
counted as unknown until you click Connect again or empty its keys.

Not covered: **personal Microsoft accounts** (Microsoft Graph has no presence for them: work or
school accounts only), Microsoft Teams free, **Google Meet** (no presence API), other people's
presence.

## 1. Microsoft Teams (work or school account)

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com/), **App registrations**,
   **New registration**: accounts in an organizational directory.
2. Platform **Web**, redirect URI: click **Connect** once on the Microsoft account of the
   integration configuration; the status at the top of the page then shows the address to register
   (usually `https://my.gladysassistant.com/redirect/oauth`). Or give an address of yours and fill
   in **Redirect URI** (see "Your own redirect URI").
3. **API permissions**: Microsoft Graph, **Delegated**: **Presence.Read** and **offline_access**.
   Your organization may require an administrator to consent.
4. **Certificates & secrets**: a new client secret. Note its end date: when it ends, the renewal is
   refused and you have to create a new one, save it and click Connect again.
5. Copy the **Application (client) ID** and the secret **value** into the configuration. Tenant:
   leave `organizations`, or give your tenant ID or domain (`contoso.com`). `common` and
   `consumers` are refused (personal accounts have no presence).
6. Save, click **Connect** and sign in on the Microsoft page.

## 2. Zoom

1. On the [Zoom App Marketplace](https://marketplace.zoom.us/), **Develop**, **Build App**: a
   **General app**, **user-managed**.
2. **OAuth redirect URL** and **allow list**: the address shown in the status after a first click on
   Connect (usually `https://my.gladysassistant.com/redirect/oauth`), or yours.
3. **Scopes**: `user:read:presence_status` (on an older app with classic scopes: `user:read`),
   nothing else.
4. Copy the **Client ID** and **Client Secret** into the configuration, save, click **Connect**,
   allow the app on the Zoom page.

Fill in only the services you use: an empty service is never read.

## Your own redirect URI

If the redirect URI of your apps is not the one of Gladys, write it in **Redirect URI**. After the
Microsoft or Zoom page, copy the whole address the browser ended on (with `code=` and `state=`) and
paste it into **Finish the authorization**, within 15 minutes. The address must be the redirect
URI given; an authorization serves once.

## Reads, pauses and stops

- One read per service per interval (60 s by default, a minute at least), whatever starts it:
  Test button, saved configuration, reconnection of Gladys, restart (the time of the last read is
  kept in `/data`).
- **429** (too many requests): paused as long as the service asks (`Retry-After`, between 1 minute
  and 6 hours; 5 minutes without it). **5xx**: next try after 2, 5, 15 then 30 minutes. Other
  errors: never sooner than the interval, then 2, 5, 15, 30 minutes. The pauses survive a restart.
- **403** on the presence (or Zoom code 5411): the permission or the account type is refused;
  next try after an hour, or right after a new Connect.
- At most 2,000 requests per service and per UTC day (a safety net; the normal pace is about 1,500).
- **Renewal refused** (secret ended, access revoked, renewal token expired): the service stops
  until you click Connect again; nothing is retried by itself, also after a restart.
- **Access refused with a new token**: the service stops; **Test** tries once (not sooner than the
  interval), or click Connect again.
- An error answer without an OAuth error code (a proxy page, a firewall) is never taken for a
  refusal: it is retried later.

## Privacy and secrets

- Your password is typed on the Microsoft or Zoom page only.
- The client secrets are kept by Gladys in its database, never shown again, sent only to the
  service that issued them.
- The access and renewal tokens are kept by Gladys in its database (in the configuration of the
  integration, never shown on the screen), sent only to the service that issued them, never in
  `/data`, the logs nor the status messages. `/data` keeps only times, pauses and stops. A
  version that kept them in `/data` moves them to Gladys at start, then removes them from `/data`.
  Zoom and Microsoft renewal tokens serve once, so a renewal is sent only while Gladys is
  connected and has just accepted a write of the current tokens; otherwise the current token stays
  and the renewal is tried again a minute later (the logs say "Gladys cannot keep the tokens"). If
  Gladys refuses the new token right after (the logs say "tokens not saved in Gladys"), it is
  saved again every 10 seconds to a minute and at each reconnection to Gladys, with no other
  renewal before. A stop of the container waits up to 8 seconds for a renewal in flight and its
  save, and tries once more to save a token Gladys refused earlier; a Connect during the stop is
  refused (click it again once the integration runs). Only a crash or a kill of the container in
  the very seconds of a renewal, or a stop while Gladys still refuses the new token, can lose it:
  the status then says the session ended; click Connect.
- The presence is kept in memory only, never in `/data` nor in the logs. "In a meeting" keeps its
  history in Gladys (you can turn it off per feature); the presence texts have none.
- The authorization uses PKCE (S256) and a state that serves once, for 15 minutes.

## Troubleshooting

- **AADSTS50020 / personal account**: use a work or school account; personal accounts have no
  presence.
- **AADSTS65001 / consent**: an administrator has to consent to Presence.Read for your
  organization.
- **invalid_client**: the client secret is wrong or ended.
- **Zoom 5411**: the app lacks the `user:read:presence_status` scope.
- The Zoom documentation reports that a user signed in on mobile only may appear `Offline`.
