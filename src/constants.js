// -----------------------------------------------------------------------------
// Manifest keys, endpoints of the two services, limits and timings.
//
// Endpoints from the official documentations only (no code reused):
//   - Microsoft identity platform, authorization code flow (v2.0 endpoints)
//     and Microsoft Graph v1.0 `GET /me/presence` (delegated permission
//     Presence.Read, work or school accounts only);
//   - Zoom OAuth (user-managed app) and Zoom API v2
//     `GET /users/me/presence_status` (scope user:read:presence_status, or
//     the classic user:read).
// -----------------------------------------------------------------------------

export const SECOND = 1000;
export const MINUTE = 60 * SECOND;
export const HOUR = 60 * MINUTE;
export const DAY = 24 * HOUR;

/** The two services, in the order of the status message. */
export const PROVIDERS = ['microsoft', 'zoom'];

/** The `oauth2` fields of the manifest `config_schema` (one Connect button each). */
export const OAUTH_FIELDS = {
  microsoft: 'microsoft_account',
  zoom: 'zoom_account',
};

export const ACTIONS = {
  FINISH_SIGN_IN: 'finish_sign_in',
  TEST_CONNECTION: 'test_connection',
};

export const MICROSOFT = {
  /** `{tenant}` is replaced by the configured tenant (organizations by default). */
  authorizeUrl: 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize',
  tokenUrl: 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token',
  presenceUrl: 'https://graph.microsoft.com/v1.0/me/presence',
  /** Presence only, and the renewal token (offline_access). Nothing else. */
  scopes: ['offline_access', 'https://graph.microsoft.com/Presence.Read'],
};

export const ZOOM = {
  authorizeUrl: 'https://zoom.us/oauth/authorize',
  tokenUrl: 'https://zoom.us/oauth/token',
  presenceUrl: 'https://api.zoom.us/v2/users/me/presence_status',
};

/** Read interval, in seconds: one read a minute at most, per account. */
export const POLL_SECONDS = { min: 60, max: 900, default: 60 };

/** What counts as "in a meeting". */
export const RULES = ['meeting', 'busy'];

/** The device type of the external ids: one device for the integration. */
export const DEVICE_TYPE = 'meeting';
export const DEVICE_PLATFORM_ID = 'status';

export const TIMINGS = {
  requestTimeout: 15 * SECOND,
  /** A button waits less: a renewal and a read stay within its 60 s. */
  quickTimeout: 8 * SECOND,
  /** The code exchange of Gladys' OAuth return: Gladys waits for its answer (5 s). */
  callbackTimeout: 3500,
  /** A pause after an ordinary error, kept in /data, read back capped at this. */
  backoffMax: HOUR,
  /** Renewed this long before its end (at most half of its lifetime). */
  tokenMargin: 10 * MINUTE,
  /** Never two renewals within a minute. */
  renewFloor: MINUTE,
  /** A renewal that did not reach the service: tried again after 1, 5, 15, then 60 minutes. */
  renewRetry: [MINUTE, 5 * MINUTE, 15 * MINUTE, HOUR],
  /** Without an expiry in the answer, the token is renewed after this. */
  defaultTokenLife: HOUR,
  /** 429: Retry-After when given, within 1 minute and 6 hours; 5 minutes without it. */
  rateLimitMin: MINUTE,
  rateLimitMax: 6 * HOUR,
  rateLimitDefault: 5 * MINUTE,
  /** 5xx: 2, 5, 15, then 30 minutes (never sooner than the interval). */
  unavailable: [2 * MINUTE, 5 * MINUTE, 15 * MINUTE, 30 * MINUTE],
  /** 403 on the presence: consent or account type; read again after 1 hour (or a new Connect). */
  forbidden: HOUR,
  /** Other failures: never sooner than the interval, then 2, 5, 15, 30 minutes. */
  backoff: [2 * MINUTE, 5 * MINUTE, 15 * MINUTE, 30 * MINUTE],
  /** A sign-in started (Connect) can be finished for 15 minutes. */
  signInTtl: 15 * MINUTE,
  testCooldown: 30 * SECOND,
  /** A reading older than max(3 intervals, this) is unknown: never published as current. */
  freshMin: 5 * MINUTE,
};

export const LIMITS = {
  /**
   * Requests sent to one service per UTC day, sign-in and renewals included:
   * 1,440 reads at the fastest pace, a renewal an hour. Reached (a loop, a
   * bug), every path of that account waits for the next UTC day.
   */
  dailyRequests: 2000,
  /** A presence answer is a few hundred bytes. */
  responseBytes: 64 * 1024,
  signIns: 5,
  /** Authorization codes, tokens: bounded before they are kept. */
  codeLength: 4096,
  tokenLength: 16384,
  pastedLength: 8192,
};
