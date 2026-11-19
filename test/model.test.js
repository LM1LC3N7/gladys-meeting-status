// -----------------------------------------------------------------------------
// The pure parts: what a presence means, the OAuth helpers, the
// configuration and the state of /data.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULTS, normalizeConfig, tenantOf } from '../src/config.js';
import { anyOf } from '../src/devices.js';
import { ServiceError } from '../src/http.js';
import { challengeOf, parseRedirect, tokenOf, unwrapState } from '../src/oauth.js';
import { inMeetingOf, presenceOf, textOf } from '../src/presence.js';
import { normalizeMemory } from '../src/store.js';

test('Teams: meetings, calls and presentations; busy only by the busy rule; unknown stays unknown', () => {
  const ms = (availability, activity, rule = 'meeting') =>
    inMeetingOf('microsoft', { availability, activity }, rule);
  assert.equal(ms('Busy', 'InAMeeting'), 1);
  assert.equal(ms('Busy', 'InACall'), 1);
  assert.equal(ms('Busy', 'InAConferenceCall'), 1);
  assert.equal(ms('DoNotDisturb', 'Presenting'), 1);
  assert.equal(ms('Available', 'Available'), 0);
  assert.equal(ms('Away', 'Away'), 0);
  assert.equal(ms('Offline', 'OffWork'), 0);
  assert.equal(ms('Busy', 'Busy'), 0);
  assert.equal(ms('Busy', 'Busy', 'busy'), 1);
  assert.equal(ms('DoNotDisturb', 'DoNotDisturb', 'busy'), 1);
  assert.equal(ms('BusyIdle', 'Inactive', 'busy'), 1);
  assert.equal(ms('Available', 'Available', 'busy'), 0);
  assert.equal(ms('PresenceUnknown', 'PresenceUnknown'), null);
  assert.equal(ms('Available', 'NewActivity'), null);
  assert.equal(ms(null, 'InAMeeting'), null);
  assert.equal(inMeetingOf('microsoft', null, 'meeting'), null);
});

test('Zoom: every documented spelling; unknown stays unknown', () => {
  const zoom = (status, rule = 'meeting') => inMeetingOf('zoom', { status }, rule);
  for (const s of [
    'In_A_Zoom_Meeting',
    'In_A_Meeting',
    'In_Meeting',
    'Presenting',
    'On_A_Call',
    'In_A_Call',
    'On_Phone_Call',
  ]) {
    assert.equal(zoom(s), 1, s);
  }
  for (const s of [
    'Available',
    'Away',
    'Offline',
    'Out_of_Office',
    'Mobile_signed_in',
    'Busy',
    'Do_Not_Disturb',
    'Do_No_Disturb',
    'In_Calendar_Event',
  ]) {
    assert.equal(zoom(s), 0, s);
  }
  for (const s of [
    'Busy',
    'Do_Not_Disturb',
    'Do_No_Disturb',
    'In_Calendar_Event',
    'In_A_Calendar_Event',
  ]) {
    assert.equal(zoom(s, 'busy'), 1, s);
  }
  assert.equal(zoom('Something_New'), null);
});

test('a presence answer: safe codes only, never a text of the service', () => {
  assert.deepEqual(
    presenceOf('microsoft', { availability: 'Busy', activity: 'InAMeeting', id: 'x' }),
    {
      availability: 'Busy',
      activity: 'InAMeeting',
    },
  );
  assert.deepEqual(presenceOf('microsoft', { availability: '<script>', activity: 'Away' }), {
    availability: null,
    activity: 'Away',
  });
  assert.throws(() => presenceOf('microsoft', {}), ServiceError);
  assert.throws(() => presenceOf('zoom', { status: 'a b' }), ServiceError);
  assert.throws(() => presenceOf('zoom', undefined), ServiceError);
  assert.equal(
    textOf('microsoft', { availability: 'Busy', activity: 'InACall' }),
    'Busy / InACall',
  );
  assert.equal(textOf('zoom', { status: 'Busy' }), 'Busy');
});

test('in a meeting, Teams or Zoom: 1 wins, 0 only when all say 0', () => {
  assert.equal(anyOf({ microsoft: 1, zoom: null }), 1);
  assert.equal(anyOf({ microsoft: 0, zoom: 1 }), 1);
  assert.equal(anyOf({ microsoft: 0, zoom: 0 }), 0);
  assert.equal(anyOf({ microsoft: 0, zoom: null }), null);
  assert.equal(anyOf({}), null);
});

test('OAuth helpers: PKCE challenge, pasted address, wrapped state, token answer', () => {
  // S256 (RFC 7636): base64url of the SHA-256 of the verifier, value from Python's hashlib.
  assert.equal(
    challengeOf('dBjftJeZ4CVP-mJ92K1QD8XOFRvDy1RW1dvbTlsbuCM'),
    'zCsQRVHALBRM-U9h_yZZjz_uOz8ArUtoNBkIivbQipY',
  );
  assert.deepEqual(parseRedirect('https://a.example/b?code=c&state=s'), {
    code: 'c',
    state: 's',
    error: null,
    base: 'https://a.example/b',
  });
  assert.equal(parseRedirect('code=c&state=s').base, null);
  assert.equal(parseRedirect('hello'), null);
  assert.equal(parseRedirect(''), null);
  const wrapped = Buffer.from(
    JSON.stringify({ v: 1, origin: 'o', path: 'p', state: 'inner' }),
  ).toString('base64url');
  assert.equal(unwrapState(wrapped), 'inner');
  assert.equal(unwrapState('inner'), null);
  assert.equal(unwrapState(Buffer.from('{"v":2,"state":"x"}').toString('base64url')), null);
  assert.equal(tokenOf({}, 0), null);
  const token = tokenOf({ access_token: 'a', refresh_token: 'r', expires_in: 3599 }, 1000);
  assert.deepEqual(token, {
    accessToken: 'a',
    refreshToken: 'r',
    expiresAt: 1000 + 3_599_000,
    obtainedAt: 1000,
  });
  assert.equal(tokenOf({ access_token: 'a', expires_in: 1e12 }, 0).expiresAt, 24 * 3600 * 1000);
});

test('the configuration: defaults, bounds, tenants, ids without spaces', () => {
  const c = normalizeConfig({});
  assert.equal(c.pollSeconds, DEFAULTS.poll_seconds);
  assert.equal(c.rule, 'meeting');
  assert.equal(c.language, 'fr');
  assert.equal(c.accounts.microsoft.client, '');
  assert.equal(c.accounts.microsoft.started, false);
  assert.equal(normalizeConfig({ poll_seconds: 5 }).pollSeconds, 60, 'one read a minute at most');
  assert.equal(normalizeConfig({ poll_seconds: 99999 }).pollSeconds, 900);
  assert.equal(normalizeConfig({ poll_seconds: '' }).pollSeconds, 60);
  assert.equal(normalizeConfig({ rule: 'other' }).rule, 'meeting');
  assert.equal(tenantOf(''), 'organizations');
  assert.equal(tenantOf('Contoso.com'), 'contoso.com');
  assert.equal(
    tenantOf('72f988bf-86f1-41af-91ab-2d7cd011db47'),
    '72f988bf-86f1-41af-91ab-2d7cd011db47',
  );
  assert.equal(tenantOf('common'), '');
  assert.equal(tenantOf('consumers'), '');
  assert.equal(tenantOf('a/b'), '');
  const z = normalizeConfig({ zoom_client_id: 'id x', zoom_client_secret: 's' }).accounts.zoom;
  assert.equal(z.clientId, '', 'a space is never part of an id');
  assert.equal(z.client, '');
  assert.equal(z.started, true);
  const crlf = normalizeConfig({ zoom_client_id: 'id', zoom_client_secret: 's\r\nX: y' }).accounts
    .zoom;
  assert.equal(crlf.clientSecret, '');
  const a = normalizeConfig({ zoom_client_id: 'id', zoom_client_secret: 's' }).accounts.zoom.client;
  const b = normalizeConfig({ zoom_client_id: 'id2', zoom_client_secret: 's' }).accounts.zoom
    .client;
  assert.match(a, /^[0-9a-f]{16}$/);
  assert.notEqual(a, b);
  const m1 = normalizeConfig({ microsoft_client_id: 'id', microsoft_client_secret: 's' });
  const m2 = normalizeConfig({
    microsoft_client_id: 'id',
    microsoft_client_secret: 's',
    microsoft_tenant: 'contoso.com',
  });
  assert.notEqual(
    m1.accounts.microsoft.client,
    m2.accounts.microsoft.client,
    'another tenant, another session',
  );
  assert.equal(normalizeConfig({ redirect_uri: 'http://example.org/x' }).redirectUri, '');
  assert.equal(
    normalizeConfig({ redirect_uri: 'http://localhost:8080/x' }).redirectUri,
    'http://localhost:8080/x',
  );
});

test('the state of /data: repaired, pauses capped, a future time taken as now, a new day restarts the budget', () => {
  const now = Date.UTC(2026, 10, 16, 12);
  const m = normalizeMemory(
    {
      microsoft: {
        session: { client: '0123456789abcdef', refreshToken: 'r', accessToken: 'a', expiresAt: 5 },
        block: { kind: 'other', client: '0123456789abcdef' },
        pause: { kind: 'rate', until: now + 30 * 24 * 3600 * 1000 },
        lastCycleAt: now + 3600 * 1000,
        requests: { day: 1, count: 1999 },
      },
      zoom: {
        session: { client: 'nothex', refreshToken: 'r' },
        pause: { kind: 'x', until: now + 1 },
      },
    },
    now,
  );
  assert.equal(m.microsoft.legacySession.refreshToken, 'r', 'tokens of an older /data, to move');
  assert.equal(m.microsoft.block, null);
  assert.equal(m.microsoft.pause.until, now + 6 * 3600 * 1000);
  assert.equal(m.microsoft.lastCycleAt, now);
  assert.equal(m.microsoft.requests.count, 0);
  assert.equal(m.zoom.legacySession, null);
  assert.equal(m.zoom.pause, null);
  assert.deepEqual(Object.keys(normalizeMemory(null, now)), ['microsoft', 'zoom']);
});
