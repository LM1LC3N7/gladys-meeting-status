// -----------------------------------------------------------------------------
// The Gladys device and its states (pure).
//
// One device for the integration, `meeting:status` (fixed: the accounts of
// the user change, a meeting light keeps its device), five read-only
// features, by role:
//   - in_meeting: Teams OR Zoom (1 when one of them says so; 0 when every
//     signed-in account says no; unknown otherwise, then not published);
//   - <service>_in_meeting: 1 / 0 for one service;
//   - <service>_presence: the codes of the service (text, no history).
// The in-meeting features are `unknown` / `binary` sensors: never a switch
// (nothing to command, and a "turn everything off" never targets them).
// -----------------------------------------------------------------------------

import {
  DEVICE_FEATURE_CATEGORIES as C,
  DEVICE_FEATURE_TYPES as TY,
} from '@gladysassistant/integration-sdk';
import { DEVICE_PLATFORM_ID, DEVICE_TYPE } from './constants.js';
import { DEVICE_NAME, FEATURE_NAMES, word } from './i18n.js';

const binary = { category: C.UNKNOWN, type: TY.SENSOR.BINARY, min: 0, max: 1, keep_history: true };
const text = { category: C.TEXT, type: TY.TEXT.TEXT, min: 0, max: 0, keep_history: false };

export const FEATURES = {
  in_meeting: binary,
  microsoft_in_meeting: binary,
  microsoft_presence: text,
  zoom_in_meeting: binary,
  zoom_presence: text,
};

export const deviceIds = (gladys) => gladys.externalIds(DEVICE_TYPE, DEVICE_PLATFORM_ID);

export function buildDevice(gladys, language) {
  const ids = deviceIds(gladys);
  return {
    name: word(DEVICE_NAME, language),
    external_id: ids.device,
    model: 'Microsoft Graph presence, Zoom presence status',
    should_poll: false,
    features: Object.entries(FEATURES).map(([role, spec]) => ({
      name: word(FEATURE_NAMES[role], language),
      external_id: ids.feature(role),
      category: spec.category,
      type: spec.type,
      min: spec.min,
      max: spec.max,
      read_only: true,
      has_feedback: false,
      keep_history: spec.keep_history,
    })),
  };
}

/**
 * `in_meeting` of all services: `readings` is { provider: 0 | 1 | null }
 * for the signed-in accounts only. 1 when one says 1; 0 when all say 0;
 * null otherwise (or without any signed-in account).
 */
export function anyOf(readings) {
  const values = Object.values(readings);
  if (values.includes(1)) {
    return 1;
  }
  return values.length > 0 && values.every((v) => v === 0) ? 0 : null;
}

/**
 * The states to publish: `fresh` is { provider: { inMeeting, text } } for
 * the services with a fresh reading, `signedIn` the services that count in
 * `in_meeting`. A null value is never published (never as 0).
 */
export function deviceStates(gladys, fresh, signedIn) {
  const ids = deviceIds(gladys);
  const states = [];
  const readings = {};
  for (const provider of signedIn) {
    readings[provider] = fresh[provider]?.inMeeting ?? null;
  }
  const any = anyOf(readings);
  if (any !== null) {
    states.push({ device_feature_external_id: ids.feature('in_meeting'), state: any });
  }
  for (const [provider, reading] of Object.entries(fresh)) {
    if (reading.inMeeting === 0 || reading.inMeeting === 1) {
      states.push({
        device_feature_external_id: ids.feature(`${provider}_in_meeting`),
        state: reading.inMeeting,
      });
    }
    if (typeof reading.text === 'string' && reading.text) {
      states.push({
        device_feature_external_id: ids.feature(`${provider}_presence`),
        text: reading.text.slice(0, 100),
      });
    }
  }
  return states;
}
