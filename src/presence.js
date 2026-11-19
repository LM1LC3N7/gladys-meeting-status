// -----------------------------------------------------------------------------
// One read of the presence of the signed-in user, and what it means.
//
// Microsoft Graph `GET /me/presence`: `availability` (Available, Busy,
// DoNotDisturb, Away…) and `activity` (InAMeeting, InACall, Presenting…).
// Zoom `GET /users/me/presence_status`: `status` (In_A_Zoom_Meeting,
// On_A_Call, Presenting, Busy, Do_Not_Disturb…; the documentation lists
// several spellings, all are known here).
//
// "In a meeting" (rule `meeting`): a meeting, a call or a presentation.
// Rule `busy`: also busy, do not disturb and, for Zoom, a calendar event.
// A value the documentation does not list, or "unknown", is UNKNOWN: never
// published as "not in a meeting".
// -----------------------------------------------------------------------------

import { ZOOM } from './constants.js';
import { endpointOf } from './oauth.js';
import { ServiceError, apiError, parseJson, safeCode } from './http.js';

export const MS_AVAILABILITY = new Set([
  'Available',
  'AvailableIdle',
  'Away',
  'BeRightBack',
  'Busy',
  'BusyIdle',
  'DoNotDisturb',
  'Offline',
]);
export const MS_ACTIVITY = new Set([
  'Available',
  'Away',
  'BeRightBack',
  'Busy',
  'DoNotDisturb',
  'InACall',
  'InAConferenceCall',
  'Inactive',
  'InAMeeting',
  'Offline',
  'OffWork',
  'OutOfOffice',
  'Presenting',
  'UrgentInterruptionsOnly',
]);
const MS_MEETING = new Set(['InACall', 'InAConferenceCall', 'InAMeeting', 'Presenting']);
const MS_BUSY_ACTIVITY = new Set(['Busy', 'DoNotDisturb', 'UrgentInterruptionsOnly']);
const MS_BUSY_AVAILABILITY = new Set(['Busy', 'BusyIdle', 'DoNotDisturb']);

export const ZOOM_STATUS = new Set([
  'Available',
  'Away',
  'Offline',
  'Mobile_signed_in',
  'Out_of_Office',
  'Busy',
  'Do_Not_Disturb',
  'Do_No_Disturb',
  'In_Calendar_Event',
  'In_A_Calendar_Event',
  'Presenting',
  'In_A_Zoom_Meeting',
  'In_A_Meeting',
  'In_Meeting',
  'On_A_Call',
  'In_A_Call',
  'On_Phone_Call',
]);
const ZOOM_MEETING = new Set([
  'Presenting',
  'In_A_Zoom_Meeting',
  'In_A_Meeting',
  'In_Meeting',
  'On_A_Call',
  'In_A_Call',
  'On_Phone_Call',
]);
const ZOOM_BUSY = new Set([
  'Busy',
  'Do_Not_Disturb',
  'Do_No_Disturb',
  'In_Calendar_Event',
  'In_A_Calendar_Event',
]);

/** The presence URL of a service. */
export const presenceUrl = (provider) =>
  provider === 'microsoft' ? endpointOf('microsoft', 'presenceUrl') : ZOOM.presenceUrl;

/**
 * The presence answer -> `{ availability, activity }` (Microsoft) or
 * `{ status }` (Zoom): safe codes only (letters, digits, `_.-`), or an
 * `api_error` when nothing usable is in it.
 */
export function presenceOf(provider, data) {
  if (provider === 'microsoft') {
    const availability = safeCode(data?.availability) ?? null;
    const activity = safeCode(data?.activity) ?? null;
    if (!availability && !activity) {
      throw new ServiceError('api_error', 'Microsoft answered without a presence');
    }
    return { availability, activity };
  }
  const status = safeCode(data?.status) ?? null;
  if (!status) {
    throw new ServiceError('api_error', 'Zoom answered without a presence');
  }
  return { status };
}

/** A presence -> 1 (in a meeting), 0 (not), or null (unknown), by `rule`. */
export function inMeetingOf(provider, presence, rule) {
  if (!presence) {
    return null;
  }
  if (provider === 'microsoft') {
    const { availability, activity } = presence;
    // Both must be values the documentation lists (PresenceUnknown is not).
    if (!MS_AVAILABILITY.has(availability) || !MS_ACTIVITY.has(activity)) {
      return null;
    }
    if (MS_MEETING.has(activity)) {
      return 1;
    }
    if (
      rule === 'busy' &&
      (MS_BUSY_ACTIVITY.has(activity) || MS_BUSY_AVAILABILITY.has(availability))
    ) {
      return 1;
    }
    return 0;
  }
  const { status } = presence;
  if (!ZOOM_STATUS.has(status)) {
    return null;
  }
  if (ZOOM_MEETING.has(status)) {
    return 1;
  }
  return rule === 'busy' && ZOOM_BUSY.has(status) ? 1 : 0;
}

/** The text of a presence for the text feature (the codes of the service). */
export function textOf(provider, presence) {
  if (!presence) {
    return null;
  }
  if (provider === 'microsoft') {
    return [presence.availability, presence.activity].filter(Boolean).join(' / ');
  }
  return presence.status;
}

/**
 * One GET of the presence with an access token. 401: `unauthorized`;
 * 403 (and Zoom's 400 code 5411, "no permission"): `forbidden`; 429, 5xx:
 * transient; anything else not 200: `http`.
 */
export async function readPresence(transport, provider, { token, now, timeout }) {
  const res = await transport(presenceUrl(provider), {
    method: 'GET',
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    timeout,
  });
  if (provider === 'zoom' && res.status === 400) {
    let data;
    try {
      data = parseJson(res.body);
    } catch {
      data = undefined;
    }
    if (Number(data?.code) === 5411) {
      throw new ServiceError('forbidden', 'the service refused the presence', {
        status: 400,
        detail: '5411',
      });
    }
  }
  const error = apiError(res, now);
  if (error) {
    throw error;
  }
  return presenceOf(provider, parseJson(res.body));
}
