// -----------------------------------------------------------------------------
// The accounts: the OAuth2 authorization code grant (with PKCE) of the app
// the user registered at Microsoft (Entra ID) or Zoom (Marketplace).
//
// The browser comes back to the redirect URI of the app:
//   - Gladys' own address (the Connect button of an `oauth2` field): Gladys
//     relays the code to onOAuthCallback;
//   - another one (configuration "redirect URI"): the user pastes the address
//     the browser ended on into the "Finish the authorization" action.
// The password is typed on the Microsoft or Zoom page only: it never reaches
// this code. Nothing here logs a token, a code or the client secret.
// -----------------------------------------------------------------------------

import { createHash, randomBytes } from 'node:crypto';
import { LIMITS, MICROSOFT, TIMINGS, ZOOM } from './constants.js';
import { ServiceError, parseJson, safeCode, transientError } from './http.js';

/** The URL of an endpoint of `provider` (Microsoft: in the configured tenant). */
export function endpointOf(provider, kind, tenant) {
  if (provider === 'microsoft') {
    return MICROSOFT[kind].replace('{tenant}', encodeURIComponent(tenant));
  }
  return ZOOM[kind];
}

/** PKCE S256: the challenge of a verifier. */
export const challengeOf = (verifier) => createHash('sha256').update(verifier).digest('base64url');

/** The authorization address of a service, for a client id, a redirect URI, a state. */
export function authorizeUrl({ provider, clientId, tenant, redirectUri, state, verifier }) {
  const query = new URLSearchParams({
    response_type: 'code',
    client_id: clientId,
    redirect_uri: redirectUri,
    state,
    code_challenge: challengeOf(verifier),
    code_challenge_method: 'S256',
  });
  if (provider === 'microsoft') {
    // The scopes are asked here; Zoom takes those set on the app itself.
    query.set('scope', MICROSOFT.scopes.join(' '));
    query.set('response_mode', 'query');
  }
  return `${endpointOf(provider, 'authorizeUrl', tenant)}?${query}`;
}

/**
 * The parameters of a pasted address: `{ code, state, error, base }`, or null
 * when nothing usable is in it. Takes the whole address (`base`: its origin
 * and path, checked against the redirect URI), or only its query string
 * ("?code=…&state=…" or "code=…&state=…", `base` null).
 */
export function parseRedirect(input) {
  const text = String(input ?? '')
    .trim()
    .slice(0, LIMITS.pastedLength);
  if (!text) {
    return null;
  }
  let query;
  let base = null;
  if (/^https?:\/\//i.test(text)) {
    let url;
    try {
      url = new URL(text.split(/\s/)[0]);
    } catch {
      return null;
    }
    query = url.search.slice(1);
    base = baseOf(url);
  } else if (/^[?]?(?:code|state|error)=/.test(text)) {
    query = text.replace(/^[?]/, '').split(/\s/)[0];
  } else {
    return null;
  }
  const params = new URLSearchParams(query);
  const code = params.get('code') || null;
  const result = {
    code: code && code.length <= LIMITS.codeLength ? code : null,
    state: params.get('state') || null,
    error: params.get('error') || null,
    base,
  };
  return result.code || result.error ? result : null;
}

/**
 * The state of the integration inside a state wrapped by the Gladys front
 * (base64url of `{ v: 1, origin, path, state }`, so the page of
 * my.gladysassistant.com can send the browser back to the instance), or
 * null when it is not a wrapped state. A pasted address carries the wrapped
 * one; Gladys' own return gives the inner one.
 */
export function unwrapState(state) {
  if (typeof state !== 'string' || state.length > 4096 || !/^[A-Za-z0-9_-]+$/.test(state)) {
    return null;
  }
  try {
    const data = JSON.parse(Buffer.from(state, 'base64url').toString('utf8'));
    return data?.v === 1 && typeof data.state === 'string' && data.state ? data.state : null;
  } catch {
    return null;
  }
}

/** The origin and path of an address (no query, no fragment), or null. */
export function baseOf(address) {
  try {
    const url = address instanceof URL ? address : new URL(address);
    return `${url.origin}${url.pathname}`;
  } catch {
    return null;
  }
}

const str = (value) =>
  typeof value === 'string' && value !== '' && value.length <= LIMITS.tokenLength
    ? value
    : undefined;

/** The token answer -> the token kept, or null without an access token. */
export function tokenOf(data, now) {
  const accessToken = str(data?.access_token);
  if (!accessToken) {
    return null;
  }
  const expiresIn = Number(data.expires_in);
  return {
    accessToken,
    refreshToken: str(data.refresh_token) ?? null,
    expiresAt:
      now +
      (Number.isFinite(expiresIn) && expiresIn > 0
        ? Math.min(expiresIn * 1000, 24 * 3600 * 1000)
        : TIMINGS.defaultTokenLife),
    obtainedAt: now,
  };
}

/** OAuth error codes that say "try later", never a refusal. */
const TRANSIENT_CODES = new Set(['temporarily_unavailable', 'server_error']);

/**
 * The token endpoint (a code or the renewal token), form-encoded.
 * Microsoft: the client id and secret in the body, the scopes again;
 * Zoom: HTTP Basic with the client id and secret. A 400 or 401 carrying an
 * OAuth `error` code: `refused` (its safe code as detail). A 400, 401 or 403
 * without one (a proxy, a firewall page): `http`, never a refusal. 429, 5xx:
 * transient; 200 without a token: `api_error`.
 */
export async function tokenRequest(
  transport,
  form,
  { provider, clientId, clientSecret, tenant, now, timeout },
) {
  const fields = { ...form };
  const headers = {
    'Content-Type': 'application/x-www-form-urlencoded',
    Accept: 'application/json',
  };
  if (provider === 'microsoft') {
    fields.client_id = clientId;
    fields.client_secret = clientSecret;
    fields.scope = MICROSOFT.scopes.join(' ');
  } else {
    headers.Authorization = `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString('base64')}`;
  }
  const body = new URLSearchParams(fields).toString();
  headers['Content-Length'] = Buffer.byteLength(body);
  const res = await transport(endpointOf(provider, 'tokenUrl', tenant), {
    method: 'POST',
    headers,
    body,
    timeout,
  });
  const transient = transientError(res, now);
  if (transient) {
    throw transient;
  }
  let data;
  try {
    data = parseJson(res.body);
  } catch (err) {
    if (res.status >= 400) {
      data = undefined;
    } else {
      throw err;
    }
  }
  const code = typeof data?.error === 'string' ? safeCode(data.error) : undefined;
  if ((res.status === 400 || res.status === 401) && code) {
    if (TRANSIENT_CODES.has(code)) {
      throw new ServiceError('unavailable', 'the sign-in service is unavailable', {
        status: res.status,
        detail: code,
      });
    }
    throw new ServiceError('refused', 'the sign-in was refused', {
      status: res.status,
      detail: code,
    });
  }
  if (res.status !== 200) {
    throw new ServiceError('http', 'unexpected answer of the sign-in', { status: res.status });
  }
  const token = tokenOf(data, now);
  if (!token) {
    throw new ServiceError('api_error', 'the sign-in answered without a token');
  }
  return token;
}

/**
 * The sign-ins started (Connect), waiting for their code:
 * state -> { provider, client, redirectUri, verifier, at }. In memory only.
 */
export class SignIns {
  constructor({ now }) {
    this.now = now;
    this.pending = new Map();
  }

  /** A new sign-in: its state, kept 15 minutes. */
  begin({ provider, client, redirectUri }) {
    const now = this.now();
    for (const [state, entry] of this.pending) {
      if (now - entry.at > TIMINGS.signInTtl || now < entry.at) {
        this.pending.delete(state);
      }
    }
    while (this.pending.size >= LIMITS.signIns) {
      this.pending.delete(this.pending.keys().next().value);
    }
    const state = randomBytes(16).toString('base64url');
    const verifier = randomBytes(32).toString('base64url');
    const entry = { provider, client, redirectUri, verifier, at: now };
    this.pending.set(state, entry);
    return { state, entry };
  }

  valid(entry) {
    const age = this.now() - entry.at;
    return age >= 0 && age <= TIMINGS.signInTtl;
  }

  /** The sign-in of `state` still valid, without using it (null when none). */
  peek(state) {
    const entry = typeof state === 'string' && state ? this.pending.get(state) : undefined;
    return entry && this.valid(entry) ? entry : null;
  }

  /** The sign-in of `state` (a state is required); it serves once. */
  take(state) {
    const entry = typeof state === 'string' && state ? this.pending.get(state) : undefined;
    if (!entry) {
      return null;
    }
    this.pending.delete(state);
    return this.valid(entry) ? { key: state, entry } : null;
  }

  /** Put a sign-in back (the service was not reached: the same code may be tried again). */
  restore({ key, entry }) {
    this.pending.set(key, entry);
  }

  /** Drops the sign-ins of one service (its keys changed), or all. */
  clear(provider) {
    for (const [state, entry] of this.pending) {
      if (!provider || entry.provider === provider) {
        this.pending.delete(state);
      }
    }
  }
}
