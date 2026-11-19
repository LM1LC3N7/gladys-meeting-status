// -----------------------------------------------------------------------------
// A fake Microsoft identity platform + Graph, and a fake Zoom, behind the
// transport of the app: the token endpoints (authorization code with PKCE,
// renewal tokens that serve once) and the presence. Every request is
// recorded; `queue(route, answer)` injects one answer
// (routes: ms_token, ms_presence, zoom_token, zoom_presence).
// -----------------------------------------------------------------------------

import { createHash } from 'node:crypto';

export const MS_CLIENT_ID = '11111111-2222-3333-4444-555555555555';
export const MS_SECRET = 'ms-secret-Very~Secret.Value';
export const ZOOM_CLIENT_ID = 'zoomClientId123';
export const ZOOM_SECRET = 'zoomSecretValue456';
export const GLADYS_REDIRECT = 'https://my.gladysassistant.com/redirect/oauth';

export const SECOND = 1000;
export const MINUTE = 60 * SECOND;
export const HOUR = 60 * MINUTE;
export const DAY = 24 * HOUR;

export const settle = async (rounds = 30) => {
  for (let i = 0; i < rounds; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
  }
};

/** /data in memory. `fail`: load throws (unreadable), `failSave`: save throws. */
export function memoryStore(data = null) {
  return {
    data,
    saves: 0,
    fail: false,
    failSave: false,
    async load() {
      if (this.fail) {
        throw Object.assign(new Error('EACCES'), { code: 'EACCES' });
      }
      return this.data === null ? null : JSON.parse(JSON.stringify(this.data));
    },
    async save(value) {
      if (this.failSave) {
        throw Object.assign(new Error('EROFS'), { code: 'EROFS' });
      }
      this.saves += 1;
      this.data = JSON.parse(JSON.stringify(value));
    },
  };
}

export const answer = (status, body, headers = {}) => ({
  status,
  headers: { 'content-type': 'application/json', ...headers },
  body: body === undefined ? '' : typeof body === 'string' ? body : JSON.stringify(body),
});

const MS_TOKEN = /^https:\/\/login\.microsoftonline\.com\/([^/]+)\/oauth2\/v2\.0\/token$/;

function provider(name, { clientId, secret, expiresIn }) {
  return {
    name,
    clientId,
    secret,
    expiresIn,
    issued: 0,
    access: null,
    refresh: null,
    used: new Set(),
    codes: new Map(),
  };
}

export function fakeServices() {
  const server = {
    requests: [],
    queued: {},
    ms: provider('ms', { clientId: MS_CLIENT_ID, secret: MS_SECRET, expiresIn: 3600 }),
    zoom: provider('zoom', { clientId: ZOOM_CLIENT_ID, secret: ZOOM_SECRET, expiresIn: 3600 }),
    /** The presence each service answers. */
    msPresence: { availability: 'Available', activity: 'Available' },
    zoomPresence: { status: 'Available' },
    queue(route, value, times = 1) {
      for (let i = 0; i < times; i += 1) {
        (this.queued[route] ??= []).push(value);
      }
    },
    of(route) {
      return this.requests.filter((r) => r.route === route);
    },
    /** The user approves on the page of `authorizeUrl`: the code and state of the return. */
    approve(authorizeUrl) {
      const url = new URL(authorizeUrl);
      const p = url.hostname === 'zoom.us' ? server.zoom : server.ms;
      const code = `${p.name}-code-${p.codes.size + 1}`;
      p.codes.set(code, {
        redirectUri: url.searchParams.get('redirect_uri'),
        clientId: url.searchParams.get('client_id'),
        challenge: url.searchParams.get('code_challenge'),
        scope: url.searchParams.get('scope'),
      });
      return {
        code,
        state: url.searchParams.get('state'),
        redirectUri: url.searchParams.get('redirect_uri'),
        url,
      };
    },
    issue(p) {
      p.issued += 1;
      p.access = `${p.name}-access-${p.issued}`;
      p.refresh = `${p.name}-refresh-${p.issued}`;
      return answer(200, {
        access_token: p.access,
        refresh_token: p.refresh,
        expires_in: p.expiresIn,
        token_type: 'Bearer',
      });
    },
    token(p, form, clientId, clientSecret) {
      if (clientId !== p.clientId || clientSecret !== p.secret) {
        return answer(401, { error: 'invalid_client' });
      }
      if (form.grant_type === 'authorization_code') {
        const entry = p.codes.get(form.code);
        p.codes.delete(form.code);
        const challenge = createHash('sha256')
          .update(form.code_verifier ?? '')
          .digest('base64url');
        if (
          !entry ||
          entry.clientId !== clientId ||
          entry.redirectUri !== form.redirect_uri ||
          entry.challenge !== challenge
        ) {
          return answer(400, { error: 'invalid_grant' });
        }
        return server.issue(p);
      }
      if (form.grant_type === 'refresh_token') {
        if (form.refresh_token !== p.refresh || p.used.has(form.refresh_token)) {
          return answer(400, { error: 'invalid_grant' });
        }
        p.used.add(form.refresh_token);
        return server.issue(p);
      }
      return answer(400, { error: 'unsupported_grant_type' });
    },
    transport: async (url, { method = 'GET', headers = {}, body, timeout } = {}) => {
      const request = { url, method, headers, body, timeout, route: undefined };
      server.requests.push(request);
      const ms = MS_TOKEN.exec(url);
      if (ms) {
        request.route = 'ms_token';
        request.tenant = decodeURIComponent(ms[1]);
        request.form = Object.fromEntries(new URLSearchParams(body));
      } else if (url === 'https://graph.microsoft.com/v1.0/me/presence') {
        request.route = 'ms_presence';
      } else if (url === 'https://zoom.us/oauth/token') {
        request.route = 'zoom_token';
        request.form = Object.fromEntries(new URLSearchParams(body));
      } else if (url === 'https://api.zoom.us/v2/users/me/presence_status') {
        request.route = 'zoom_presence';
      } else {
        return answer(404, { error: 'not found' });
      }
      const injected = server.queued[request.route]?.shift();
      if (injected) {
        return typeof injected === 'function' ? injected(request) : injected;
      }
      switch (request.route) {
        case 'ms_token':
          return server.token(
            server.ms,
            request.form,
            request.form.client_id,
            request.form.client_secret,
          );
        case 'zoom_token': {
          const basic = /^Basic (.+)$/.exec(headers.Authorization ?? '');
          const [id, secret] = basic
            ? Buffer.from(basic[1], 'base64').toString('utf8').split(':')
            : [];
          return server.token(server.zoom, request.form, id, secret);
        }
        case 'ms_presence':
          if (!server.ms.access || headers.Authorization !== `Bearer ${server.ms.access}`) {
            return answer(401, { error: { code: 'InvalidAuthenticationToken' } });
          }
          return answer(200, { id: 'user-guid', ...server.msPresence });
        default:
          if (!server.zoom.access || headers.Authorization !== `Bearer ${server.zoom.access}`) {
            return answer(401, { code: 124, message: 'Invalid access token.' });
          }
          return answer(200, server.zoomPresence);
      }
    },
  };
  return server;
}
