// -----------------------------------------------------------------------------
// The configuration of the manifest `config_schema`, normalized.
// -----------------------------------------------------------------------------

import { createHash } from 'node:crypto';
import { POLL_SECONDS, RULES } from './constants.js';
import { sessionOf } from './store.js';

/** Defaults, mirrored by the manifest (checked by test/manifest.test.js). */
export const DEFAULTS = {
  microsoft_tenant: 'organizations',
  poll_seconds: POLL_SECONDS.default,
  rule: 'meeting',
  language: 'fr',
};

export const LANGUAGES = ['fr', 'en'];

/**
 * A short hash of what a session belongs to (service, client id, tenant):
 * the session in /data is dropped when one of them changes.
 */
export const clientKey = (...parts) =>
  createHash('sha256')
    .update(`gladys-meeting-status:client:${parts.join('\n')}`)
    .digest('hex')
    .slice(0, 16);

const bounded = (raw, { min, max, default: fallback }) => {
  const value = Number(raw);
  return raw === null || raw === undefined || raw === '' || !Number.isFinite(value)
    ? fallback
    : Math.round(Math.max(min, Math.min(max, value)));
};

/** A short text without spaces nor control characters (an id, a secret), or ''. */
const token = (raw, max = 256) => {
  const text = String(raw ?? '').trim();
  // eslint-disable-next-line no-control-regex
  return text && text.length <= max && !/[\s\u0000-\u001f\u007f]/.test(text) ? text : '';
};

/**
 * The Microsoft tenant: `organizations` (any work or school account),
 * a tenant id (GUID) or a verified domain. `common` and `consumers` would
 * let a personal account sign in, which has no presence: refused.
 */
export function tenantOf(raw) {
  const text = String(raw ?? '')
    .trim()
    .toLowerCase();
  if (!text) {
    return DEFAULTS.microsoft_tenant;
  }
  if (text === 'organizations') {
    return text;
  }
  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(text)) {
    return text;
  }
  if (
    text.length <= 253 &&
    /^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/.test(text) &&
    text !== 'common' &&
    text !== 'consumers'
  ) {
    return text;
  }
  return '';
}

/** An http(s) address without spaces, or '' (a wrong one is ignored). */
function address(raw) {
  const text = String(raw ?? '').trim();
  if (!text || text.length > 2048 || /\s/.test(text)) {
    return '';
  }
  try {
    const url = new URL(text);
    // HTTPS only, except a local address (the code never crosses a network in clear).
    const local = ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
    return url.protocol === 'https:' || (url.protocol === 'http:' && local) ? url.toString() : '';
  } catch {
    return '';
  }
}

/** The key of the configuration of Gladys holding the tokens of a service (outside the schema). */
export const sessionKey = (provider) => `${provider}_session`;

function account(provider, clientId, clientSecret, tenant = '', stored = null) {
  return {
    // The tokens saved in Gladys (read once, at start).
    stored,
    provider,
    clientId,
    clientSecret,
    tenant,
    // Computed from the values kept, never from a refused one; '' when incomplete.
    client:
      clientId && clientSecret && (provider !== 'microsoft' || tenant)
        ? clientKey(provider, clientId, tenant)
        : '',
    // Something typed: an incomplete account says what is missing.
    started: Boolean(clientId || clientSecret),
  };
}

export function normalizeConfig(raw = {}) {
  const language = String(raw?.language ?? DEFAULTS.language)
    .trim()
    .toLowerCase();
  const rule = String(raw?.rule ?? DEFAULTS.rule)
    .trim()
    .toLowerCase();
  return {
    accounts: {
      microsoft: account(
        'microsoft',
        token(raw?.microsoft_client_id),
        token(raw?.microsoft_client_secret, 1024),
        tenantOf(raw?.microsoft_tenant),
        sessionOf(raw?.[sessionKey('microsoft')]),
      ),
      zoom: account(
        'zoom',
        token(raw?.zoom_client_id),
        token(raw?.zoom_client_secret, 1024),
        '',
        sessionOf(raw?.[sessionKey('zoom')]),
      ),
    },
    redirectUri: address(raw?.redirect_uri),
    pollSeconds: bounded(raw?.poll_seconds, POLL_SECONDS),
    rule: RULES.includes(rule) ? rule : DEFAULTS.rule,
    language: LANGUAGES.includes(language) ? language : DEFAULTS.language,
  };
}
