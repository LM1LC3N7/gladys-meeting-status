// -----------------------------------------------------------------------------
// End to end: the real SDK wiring (src/handlers.js) against the SDK's fake
// Gladys and fake Microsoft and Zoom services behind the transport. The
// clock is the kit's test clock: it moves at every read, timers fire only
// when the test advances it.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DEVICE_TRANSPORTS } from '@gladysassistant/integration-sdk';
import { createFakeGladys } from '@gladysassistant/integration-sdk/testing';
import { MeetingApp } from '../src/app.js';
import { clientKey } from '../src/config.js';
import { ACTIONS, LIMITS, OAUTH_FIELDS } from '../src/constants.js';
import { registerHandlers } from '../src/handlers.js';
import { testClock } from './kit/index.mjs';
import {
  DAY,
  GLADYS_REDIRECT,
  HOUR,
  MINUTE,
  MS_CLIENT_ID,
  MS_SECRET,
  SECOND,
  ZOOM_CLIENT_ID,
  ZOOM_SECRET,
  answer,
  fakeServices,
  memoryStore,
  settle,
} from './helpers/fakeServices.js';

const START = Date.UTC(2026, 10, 16, 9, 0, 0);
const MS = { microsoft_client_id: MS_CLIENT_ID, microsoft_client_secret: MS_SECRET };
const ZOOM = { zoom_client_id: ZOOM_CLIENT_ID, zoom_client_secret: ZOOM_SECRET };
const BOTH = { ...MS, ...ZOOM };
const MS_CLIENT = clientKey('microsoft', MS_CLIENT_ID, 'organizations');
const DEVICE = 'ext:test-integration:meeting:status';
const feature = (key) => `${DEVICE}:${key}`;

const recorder = (logs) =>
  Object.fromEntries(
    ['debug', 'info', 'warn', 'error'].map((level) => [
      level,
      (...args) => logs.push(args.map((a) => (a instanceof Error ? a.stack : String(a))).join(' ')),
    ]),
  );

async function setup(
  t,
  {
    config = MS,
    store = memoryStore(),
    server = fakeServices(),
    clock = testClock({ start: START }),
    signIn = store.data === null
      ? ['microsoft', 'zoom'].filter((p) => config[`${p}_client_id`])
      : [],
    create = true,
    logs = [],
    devices,
  } = {},
) {
  const gladys = createFakeGladys({ config, ...(devices ? { devices } : {}) });
  const app = new MeetingApp({
    gladys,
    logger: recorder(logs),
    transport: server.transport,
    now: clock.now,
    mono: clock.monotonic,
    timers: clock.timers,
    store,
  });
  registerHandlers(gladys, app);
  t.after(() => app.shutdown());
  await gladys.connect();
  await settle();
  const advance = async (ms, step = 5 * SECOND) => {
    for (let elapsed = 0; elapsed < ms; elapsed += step) {
      await clock.advance(Math.min(step, ms - elapsed));
      await settle();
    }
  };
  const connect = async (provider, redirect = GLADYS_REDIRECT) => {
    const ack = await gladys.fake.oauthAuthorizeUrl(OAUTH_FIELDS[provider], redirect);
    assert.equal(ack.success, true, JSON.stringify(ack));
    return server.approve(ack.data.authorize_url);
  };
  const authorize = async (provider) => {
    const back = await connect(provider);
    const ack = await gladys.fake.oauthCallback(OAUTH_FIELDS[provider], back);
    assert.equal(ack.success, true, JSON.stringify(ack));
    await settle();
  };
  for (const provider of signIn) {
    await authorize(provider);
  }
  if (create) {
    for (const d of gladys.fake.discoveredDevices) {
      await gladys.fake.deviceCreated(d);
    }
    await settle();
  }
  return { gladys, app, server, store, clock, advance, connect, authorize, logs };
}

const message = (gladys) => gladys.fake.connectionStatus?.message ?? {};
const last = (gladys, key) => gladys.fake.lastState(feature(key));
const statesOf = (gladys, key) =>
  gladys.fake.states
    .filter((s) => s.device_feature_external_id === feature(key))
    .map((s) => s.text ?? s.state);

test('nothing configured: the status asks for an app, nothing is sent, Connect refuses', async (t) => {
  const { gladys, server, advance } = await setup(t, { config: {} });
  assert.equal(gladys.fake.connectionStatus.connected, false);
  assert.match(message(gladys).en, /client id and the client secret of your Microsoft/);
  for (const field of Object.values(OAUTH_FIELDS)) {
    const ack = await gladys.fake.oauthAuthorizeUrl(field, GLADYS_REDIRECT);
    assert.equal(ack.success, false);
  }
  await advance(2 * HOUR, 10 * MINUTE);
  assert.equal(server.requests.length, 0);
  assert.equal(gladys.fake.discoveredDevices.length, 0, 'no device before an authorization');
});

test('an incomplete account says what is missing; a personal-account tenant is refused', async (t) => {
  const { gladys } = await setup(t, {
    config: { ...MS, microsoft_tenant: 'common', zoom_client_id: ZOOM_CLIENT_ID },
    signIn: [],
  });
  assert.match(message(gladys).en, /Microsoft Teams: the client id, the client secret/);
  assert.match(message(gladys).en, /Zoom: the client id, the client secret/);
  assert.match(message(gladys).fr, /Microsoft Teams : /);
});

test('signed out: never a request by itself; Connect builds the Microsoft address with PKCE', async (t) => {
  const { gladys, server, advance, connect } = await setup(t, { signIn: [] });
  assert.match(message(gladys).en, /Microsoft Teams: not signed in: click Connect/);
  await advance(2 * HOUR, 10 * MINUTE);
  assert.equal(server.requests.length, 0);
  const ack = await gladys.fake.oauthAuthorizeUrl(OAUTH_FIELDS.microsoft, GLADYS_REDIRECT);
  const url = new URL(ack.data.authorize_url);
  assert.equal(
    `${url.origin}${url.pathname}`,
    'https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize',
  );
  assert.equal(url.searchParams.get('response_type'), 'code');
  assert.equal(url.searchParams.get('client_id'), MS_CLIENT_ID);
  assert.equal(url.searchParams.get('redirect_uri'), GLADYS_REDIRECT);
  assert.equal(
    url.searchParams.get('scope'),
    'offline_access https://graph.microsoft.com/Presence.Read',
  );
  assert.equal(url.searchParams.get('code_challenge_method'), 'S256');
  assert.match(url.searchParams.get('code_challenge'), /^[A-Za-z0-9_-]{43}$/);
  assert.ok(url.searchParams.get('state').length >= 16);
  assert.ok(!ack.data.authorize_url.includes(MS_SECRET), 'the secret never goes to the browser');
  // Not connected: the status shows the return address to register in the app.
  assert.match(message(gladys).en, /Redirect URI to register in your app: https:\/\/my\.gladys/);
  await connect('microsoft');
  assert.equal(server.requests.length, 0);
});

test('Teams: authorization, one read, device and states; tokens only in /data, no secret logged', async (t) => {
  const logs = [];
  const { gladys, store, server } = await setup(t, { logs });
  const [exchange] = server.of('ms_token');
  assert.equal(exchange.tenant, 'organizations');
  assert.equal(exchange.form.grant_type, 'authorization_code');
  assert.equal(exchange.form.client_id, MS_CLIENT_ID);
  assert.equal(exchange.form.client_secret, MS_SECRET);
  assert.equal(exchange.form.redirect_uri, GLADYS_REDIRECT);
  assert.match(exchange.form.code_verifier, /^[A-Za-z0-9_-]{43}$/);
  assert.match(exchange.form.scope, /Presence\.Read/);
  const [read] = server.of('ms_presence');
  assert.equal(read.method, 'GET');
  assert.equal(read.headers.Authorization, `Bearer ${server.ms.access}`);
  assert.equal(read.headers.Accept, 'application/json');
  assert.equal(gladys.fake.connectionStatus.connected, true);
  assert.equal(message(gladys).en, 'Microsoft Teams: connected.');

  const [d] = gladys.fake.discoveredDevices;
  assert.equal(d.external_id, DEVICE);
  assert.equal(d.name, 'Statut en réunion');
  assert.ok(d.features.every((f) => f.read_only === true));
  assert.ok(d.features.every((f) => !['switch', 'light'].includes(f.category)));
  const binary = d.features.find((f) => f.external_id === feature('in_meeting'));
  assert.deepEqual(
    [binary.category, binary.type, binary.min, binary.max],
    ['unknown', 'binary', 0, 1],
  );

  assert.equal(last(gladys, 'in_meeting'), 0);
  assert.equal(last(gladys, 'microsoft_in_meeting'), 0);
  assert.equal(last(gladys, 'microsoft_presence'), 'Available / Available');
  assert.equal(last(gladys, 'zoom_in_meeting'), undefined);

  const kept = store.data.microsoft.session;
  assert.equal(kept.client, MS_CLIENT);
  assert.equal(kept.accessToken, server.ms.access);
  assert.equal(kept.refreshToken, server.ms.refresh);
  const everything = [
    ...logs,
    JSON.stringify(gladys.fake.connectionStatuses),
    JSON.stringify(gladys.fake.states),
  ].join('\n');
  for (const secret of [MS_SECRET, server.ms.access, server.ms.refresh, 'ms-code-1']) {
    assert.ok(!everything.includes(secret), `${secret} leaked`);
  }
  assert.ok(!JSON.stringify(store.data).includes(MS_SECRET), 'the client secret is not in /data');
  assert.ok(!JSON.stringify(store.data).includes('Available'), 'no presence in /data');
});

test('a meeting starts and ends: one read a minute, never more, states only when they change', async (t) => {
  const { gladys, server, advance } = await setup(t);
  server.msPresence = { availability: 'Busy', activity: 'InAMeeting' };
  await advance(MINUTE);
  assert.equal(last(gladys, 'in_meeting'), 1);
  assert.equal(last(gladys, 'microsoft_presence'), 'Busy / InAMeeting');
  server.msPresence = { availability: 'DoNotDisturb', activity: 'Presenting' };
  await advance(MINUTE);
  assert.equal(last(gladys, 'in_meeting'), 1);
  server.msPresence = { availability: 'Available', activity: 'Available' };
  await advance(MINUTE);
  assert.equal(last(gladys, 'in_meeting'), 0);
  assert.deepEqual(statesOf(gladys, 'in_meeting'), [0, 1, 0], 'unchanged values are not resent');
  const before = server.of('ms_presence').length;
  await advance(10 * MINUTE, SECOND);
  const reads = server.of('ms_presence').length - before;
  assert.ok(reads >= 9 && reads <= 10, `${reads} reads in 10 minutes`);
});

test('the rule: busy counts only when asked; a change applies without a read', async (t) => {
  const { gladys, server, advance } = await setup(t);
  server.msPresence = { availability: 'Busy', activity: 'Busy' };
  await advance(MINUTE);
  assert.equal(last(gladys, 'microsoft_in_meeting'), 0);
  const reads = server.of('ms_presence').length;
  await gladys.fake.configUpdated({ ...MS, rule: 'busy' });
  await settle();
  assert.equal(last(gladys, 'microsoft_in_meeting'), 1);
  assert.equal(last(gladys, 'in_meeting'), 1);
  assert.equal(server.of('ms_presence').length, reads, 'no read for a new rule');
});

test('an unknown presence is never published as "not in a meeting"', async (t) => {
  const { gladys, server, advance } = await setup(t);
  server.msPresence = { availability: 'Busy', activity: 'InAMeeting' };
  await advance(MINUTE);
  server.msPresence = { availability: 'PresenceUnknown', activity: 'PresenceUnknown' };
  await advance(MINUTE);
  assert.equal(last(gladys, 'in_meeting'), 1);
  assert.equal(last(gladys, 'microsoft_in_meeting'), 1);
  assert.equal(last(gladys, 'microsoft_presence'), 'PresenceUnknown / PresenceUnknown');
  assert.match(message(gladys).en, /presence unknown/);
  server.msPresence = { availability: 'Away', activity: 'SomethingNew' };
  await advance(MINUTE);
  assert.equal(last(gladys, 'microsoft_in_meeting'), 1, 'an activity not documented is unknown');
});

test('Teams and Zoom: Zoom by Basic auth; "in a meeting" is one OR the other', async (t) => {
  const { gladys, server, advance } = await setup(t, { config: BOTH });
  const [zoomExchange] = server.of('zoom_token');
  const basic = Buffer.from(`${ZOOM_CLIENT_ID}:${ZOOM_SECRET}`).toString('base64');
  assert.equal(zoomExchange.headers.Authorization, `Basic ${basic}`);
  assert.ok(!zoomExchange.body.includes(ZOOM_SECRET), 'the Zoom secret is not in the body');
  assert.equal(zoomExchange.form.client_id, undefined);
  assert.match(message(gladys).en, /^Microsoft Teams: connected\. Zoom: connected\.$/);
  assert.equal(last(gladys, 'in_meeting'), 0);
  server.zoomPresence = { status: 'In_A_Zoom_Meeting' };
  await advance(MINUTE);
  assert.equal(last(gladys, 'zoom_in_meeting'), 1);
  assert.equal(last(gladys, 'microsoft_in_meeting'), 0);
  assert.equal(last(gladys, 'in_meeting'), 1);
  assert.equal(last(gladys, 'zoom_presence'), 'In_A_Zoom_Meeting');
  server.zoomPresence = { status: 'Do_Not_Disturb', end_time: 'x', remaining_time: 60 };
  await advance(MINUTE);
  assert.equal(last(gladys, 'in_meeting'), 0);
});

test('a service failing for long: its value turns unknown, "in a meeting" is never set to 0 by the other', async (t) => {
  const { gladys, server, advance } = await setup(t, { config: BOTH });
  server.zoomPresence = { status: 'On_A_Call' };
  await advance(MINUTE);
  assert.equal(last(gladys, 'in_meeting'), 1);
  const count = statesOf(gladys, 'in_meeting').length;
  server.queue('zoom_presence', answer(503, ''), 100);
  await advance(2 * HOUR, 30 * SECOND);
  assert.equal(statesOf(gladys, 'in_meeting').length, count, 'Teams 0 + Zoom unknown: nothing');
  assert.match(message(gladys).en, /Zoom: the service is unavailable/);
  assert.equal(gladys.fake.connectionStatus.connected, false);
  // The device badge stays "cloud": Teams is read.
  assert.equal(gladys.fake.transports.at(-1).transport, DEVICE_TRANSPORTS.CLOUD);
});

test('5xx: growing pauses, never sooner than the interval, never a storm', async (t) => {
  const { server, advance } = await setup(t);
  server.queue('ms_presence', answer(500, ''), 100);
  const before = server.of('ms_presence').length;
  await advance(HOUR, 10 * SECOND);
  const reads = server.of('ms_presence').length - before;
  // 2, 5, 15, 30 minutes apart.
  assert.ok(reads >= 3 && reads <= 5, `${reads} reads in an hour of 5xx`);
});

test('the token is renewed before its end; the new renewal token is kept, the old one never sent again', async (t) => {
  const { store, server, advance } = await setup(t);
  const first = server.ms.refresh;
  await advance(55 * MINUTE, 30 * SECOND);
  const renewals = server.of('ms_token').filter((r) => r.form.grant_type === 'refresh_token');
  assert.equal(renewals.length, 1);
  assert.equal(renewals[0].form.refresh_token, first);
  assert.match(renewals[0].form.scope, /offline_access/);
  assert.equal(store.data.microsoft.session.refreshToken, server.ms.refresh);
  assert.notEqual(server.ms.refresh, first);
  await advance(2 * HOUR, MINUTE);
  const sent = server.of('ms_token').map((r) => r.form.refresh_token);
  assert.equal(sent.filter((x) => x === first).length, 1);
});

test('a refused renewal stops the account, kept in /data; a restart does not retry; Connect lifts it', async (t) => {
  const { app, gladys, store, server, advance } = await setup(t);
  server.queue(
    'ms_token',
    answer(400, { error: 'invalid_grant', error_description: 'AADSTS70008' }),
  );
  await advance(55 * MINUTE, 30 * SECOND);
  assert.match(message(gladys).en, /session ended \(renewal refused\)/);
  assert.equal(store.data.microsoft.session, null);
  assert.equal(store.data.microsoft.block.kind, 'expired');
  const count = server.requests.length;
  await advance(3 * HOUR, 10 * MINUTE);
  assert.equal(server.requests.length, count, 'never retried by itself');
  // A restart on the same /data: still stopped, nothing sent.
  app.stop();
  const again = await setup(t, { store, server, signIn: [] });
  await again.advance(HOUR, 10 * MINUTE);
  assert.equal(server.requests.length, count);
  assert.match(message(again.gladys).en, /session ended/);
  await again.authorize('microsoft');
  await settle();
  assert.equal(store.data.microsoft.block, null);
  assert.equal(again.gladys.fake.connectionStatus.connected, true);
  assert.equal(gladys.fake.connectionStatus.connected, false);
});

test('a token endpoint answer without an OAuth error is not a refusal (proxy page, 403)', async (t) => {
  const { gladys, store, server, advance } = await setup(t);
  server.queue('ms_token', answer(400, '<html>proxy</html>', { 'content-type': 'text/html' }));
  server.queue('ms_token', answer(403, ''));
  await advance(70 * MINUTE, 30 * SECOND);
  assert.equal(store.data.microsoft.block, null);
  assert.ok(store.data.microsoft.session, 'the session is kept');
  await advance(2 * HOUR, MINUTE);
  assert.equal(gladys.fake.connectionStatus.connected, true, 'renewed later');
});

test('401 with a token just issued: a stop (no renewal loop); the Test button tries once when due', async (t) => {
  const { gladys, store, server, advance } = await setup(t);
  // The token of the read is revoked: renewed once, the new one refused too.
  server.queue('ms_presence', answer(401, {}), 2);
  await advance(MINUTE);
  const renewals = server.of('ms_token').filter((r) => r.form.grant_type === 'refresh_token');
  assert.equal(renewals.length, 1, 'one renewal, never a loop');
  assert.equal(store.data.microsoft.block.kind, 'unauthorized');
  assert.match(message(gladys).en, /refused even with a new token/);
  const count = server.requests.length;
  await advance(HOUR, 5 * MINUTE);
  assert.equal(server.requests.length, count);
  const ack = await gladys.fake.action(ACTIONS.TEST_CONNECTION);
  assert.equal(ack.success, true);
  await settle();
  assert.equal(store.data.microsoft.block, null);
  assert.equal(gladys.fake.connectionStatus.connected, true);
});

test('429: Retry-After honored (an HTTP date against the Date header), kept in /data across a restart', async (t) => {
  const { app, gladys, store, server, clock, advance } = await setup(t);
  const date = new Date(START + 7 * DAY); // the service's clock is far from ours
  server.queue(
    'ms_presence',
    answer(
      429,
      {},
      {
        date: date.toUTCString(),
        'retry-after': new Date(date.getTime() + 600 * SECOND).toUTCString(),
      },
    ),
  );
  await advance(MINUTE);
  assert.match(message(gladys).en, /too many requests: paused for 10 min/);
  const until = store.data.microsoft.pause.until;
  assert.ok(until - clock.now() > 9 * MINUTE && until - clock.now() <= 10 * MINUTE);
  const count = server.of('ms_presence').length;
  app.stop();
  const again = await setup(t, { store, server, clock, signIn: [] });
  await again.advance(9 * MINUTE, 30 * SECOND);
  assert.equal(server.of('ms_presence').length, count, 'the pause survives a restart');
  await again.advance(2 * MINUTE, 30 * SECOND);
  assert.ok(server.of('ms_presence').length > count);
});

test('403 on the presence: an hour of pause with its reason; a new Connect lifts it', async (t) => {
  const { gladys, server, advance, authorize } = await setup(t);
  server.queue('ms_presence', answer(403, { error: { code: 'Forbidden' } }));
  await advance(MINUTE);
  assert.match(message(gladys).en, /presence was refused \(work or school account/);
  const count = server.of('ms_presence').length;
  await advance(50 * MINUTE, 5 * MINUTE);
  assert.equal(server.of('ms_presence').length, count);
  await authorize('microsoft');
  await advance(2 * MINUTE);
  assert.ok(server.of('ms_presence').length > count);
});

test('Zoom 400 code 5411 is a refused presence, not a refused sign-in', async (t) => {
  const { gladys, store, server, advance } = await setup(t, { config: ZOOM });
  server.queue('zoom_presence', answer(400, { code: 5411, message: 'no permission' }));
  await advance(MINUTE);
  assert.match(message(gladys).en, /Zoom: the presence was refused/);
  assert.equal(store.data.zoom.block, null);
});

test('a restart never reads sooner than the interval, and stays connected after a good read', async (t) => {
  const { app, store, server, clock } = await setup(t);
  const count = server.of('ms_presence').length;
  await clock.advance(20 * SECOND);
  app.stop();
  const again = await setup(t, { store, server, clock, signIn: [] });
  assert.equal(server.of('ms_presence').length, count);
  assert.match(message(again.gladys).en, /connected\. Next read in 4\d s/);
  assert.equal(again.gladys.fake.connectionStatus.connected, true);
  await again.advance(45 * SECOND, 5 * SECOND);
  assert.equal(server.of('ms_presence').length, count + 1);
});

test('getConfig() failing at connection: retried, spaced, with its status, then started', async (t) => {
  const clock = testClock({ start: START });
  const store = memoryStore();
  const server = fakeServices();
  const gladys = createFakeGladys({ config: MS });
  const app = new MeetingApp({
    gladys,
    logger: recorder([]),
    transport: server.transport,
    now: clock.now,
    mono: clock.monotonic,
    timers: clock.timers,
    store,
  });
  t.after(() => app.shutdown());
  let calls = 0;
  const statuses = [];
  gladys.setConnectionStatus = async (connected, msg) => statuses.push([connected, msg]);
  const started = app.onConnected(async () => {
    calls += 1;
    if (calls <= 2) {
      throw new Error('503');
    }
    return MS;
  });
  await settle();
  assert.match(statuses.at(-1)[1].en, /configuration could not be read.*5 s/);
  await clock.advance(5 * SECOND);
  await settle();
  assert.equal(calls, 2);
  assert.match(statuses.at(-1)[1].en, /10 s/);
  await clock.advance(10 * SECOND);
  await settle();
  await started;
  assert.equal(calls, 3);
  assert.match(statuses.at(-1)[1].en, /Microsoft Teams: not signed in/);
});

test('/data unreadable: nothing sent nor written, read again every 10 s, then the session is used', async (t) => {
  const first = await setup(t);
  const store = first.store;
  const server = first.server;
  first.app.stop();
  store.fail = true;
  const saves = store.saves;
  const count = server.requests.length;
  const again = await setup(t, { store, server, signIn: [] });
  assert.match(message(again.gladys).en, /state file of \/data cannot be read/);
  const ack = await again.gladys.fake.oauthAuthorizeUrl(OAUTH_FIELDS.microsoft, GLADYS_REDIRECT);
  assert.equal(ack.success, false);
  const test1 = await again.gladys.fake.action(ACTIONS.TEST_CONNECTION);
  assert.match(JSON.stringify(test1.data), /cannot be read/);
  await again.advance(MINUTE, SECOND);
  assert.equal(server.requests.length, count);
  assert.equal(store.saves, saves, 'never written over an unreadable file');
  store.fail = false;
  await again.advance(15 * SECOND, SECOND);
  assert.ok(server.requests.length > count, 'the session of /data is used once readable');
  assert.equal(again.gladys.fake.connectionStatus.connected, true);
});

test('a state serves once, for the service of its Connect button only', async (t) => {
  const { gladys, server, connect } = await setup(t, { config: BOTH, signIn: [] });
  const back2 = await connect('microsoft');
  const wrong = await gladys.fake.oauthCallback(OAUTH_FIELDS.zoom, back2);
  assert.equal(wrong.success, false);
  assert.match(wrong.error, /not started here/);
  assert.equal(server.requests.length, 0, 'nothing sent for the wrong service');
  // The right button still finishes it.
  assert.equal((await gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, back2)).success, true);
  const replay = await gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, back2);
  assert.equal(replay.success, false);
  const forged = await gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, {
    code: 'x',
    state: 'forged',
  });
  assert.equal(forged.success, false);
  assert.equal(server.of('ms_token').length, 1);
});

test('a refused code: a clear message with its safe code only', async (t) => {
  const { gladys, server, connect } = await setup(t, { signIn: [] });
  const back = await connect('microsoft');
  server.queue('ms_token', answer(401, { error: 'invalid_client', error_description: MS_SECRET }));
  const ack = await gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, back);
  assert.equal(ack.success, false);
  assert.match(ack.error, /refused the authorization \(invalid_client\)/);
  assert.ok(!ack.error.includes(MS_SECRET));
});

test('own redirect URI: the pasted address (state wrapped by the Gladys front) finishes the sign-in', async (t) => {
  const own = 'https://example.org/back';
  const { gladys, server } = await setup(t, {
    config: { ...ZOOM, redirect_uri: own },
    signIn: [],
  });
  const ack = await gladys.fake.oauthAuthorizeUrl(OAUTH_FIELDS.zoom, GLADYS_REDIRECT);
  const authorizeUrl = new URL(ack.data.authorize_url);
  assert.equal(authorizeUrl.searchParams.get('redirect_uri'), own);
  // The front wraps the state with its address before opening the page.
  const inner = authorizeUrl.searchParams.get('state');
  const wrapped = Buffer.from(
    JSON.stringify({ v: 1, origin: 'http://gladys.local', path: '/x', state: inner }),
  ).toString('base64url');
  authorizeUrl.searchParams.set('state', wrapped);
  const back = server.approve(authorizeUrl.toString());
  const wrong = await gladys.fake.action(ACTIONS.FINISH_SIGN_IN, {
    url: `https://evil.example/back?code=${back.code}&state=${wrapped}`,
  });
  assert.match(wrong.data.message?.en ?? JSON.stringify(wrong.data), /not the redirect URI/);
  const cancelled = await gladys.fake.action(ACTIONS.FINISH_SIGN_IN, {
    url: `${own}?error=access_denied&state=${wrapped}`,
  });
  assert.match(JSON.stringify(cancelled.data), /access_denied/);
  const done = await gladys.fake.action(ACTIONS.FINISH_SIGN_IN, {
    url: `${own}?code=${back.code}&state=${wrapped}`,
  });
  assert.match(JSON.stringify(done.data), /Zoom: Authorized/);
  await settle();
  assert.equal(server.of('zoom_presence').length, 1);
  assert.equal(server.of('zoom_token')[0].form.redirect_uri, own);
});

test('new keys for one service drop its session, never the other one', async (t) => {
  const { gladys, server, store } = await setup(t, { config: BOTH });
  await gladys.fake.configUpdated({
    ...BOTH,
    microsoft_client_id: 'aaaaaaaa-2222-3333-4444-555555555555',
  });
  await settle();
  assert.match(message(gladys).en, /Microsoft Teams: not signed in.*Zoom: connected/);
  assert.ok(store.data.zoom.session);
  const count = server.of('ms_presence').length;
  await gladys.fake.configUpdated({ ...BOTH });
  await settle();
  // Back to the first app: its session (still in /data) is used, after the interval.
  assert.equal(server.of('ms_presence').length, count);
  assert.match(message(gladys).en, /Microsoft Teams: connected/);
});

test('the daily budget: reached, every path of that account waits for the next UTC day', async (t) => {
  const first = await setup(t);
  const store = first.store;
  store.data.microsoft.requests.count = LIMITS.dailyRequests;
  store.data.microsoft.lastCycleAt = 0;
  first.app.stop();
  const count = first.server.requests.length;
  const again = await setup(t, {
    store,
    server: first.server,
    clock: testClock({ start: START + 2 * MINUTE }),
    signIn: [],
  });
  await again.advance(MINUTE);
  assert.equal(first.server.requests.length, count);
  assert.match(message(again.gladys).en, /daily number of requests reached/);
});

test('the Test button: never a read sooner than the interval, a cooldown of 30 s', async (t) => {
  const { gladys, server } = await setup(t);
  const count = server.requests.length;
  const ack = await gladys.fake.action(ACTIONS.TEST_CONNECTION);
  assert.match(JSON.stringify(ack.data), /Next read in/);
  assert.equal(server.requests.length, count);
  const again = await gladys.fake.action(ACTIONS.TEST_CONNECTION);
  assert.match(JSON.stringify(again.data), /Next read in/);
  assert.equal(server.requests.length, count);
});

test('a reconnection to Gladys resends status and Discovery, never a read', async (t) => {
  const { gladys, app, server } = await setup(t);
  const count = server.requests.length;
  const published = gladys.fake.discoveredDevices.length;
  await app.connected(MS);
  await settle();
  assert.equal(server.requests.length, count);
  assert.ok(gladys.fake.connectionStatuses.length > 0);
  assert.ok(gladys.fake.discoveredDevices.length >= published);
});

test('a device created again gets the fresh values only, never a stale one', async (t) => {
  const { gladys, server, advance } = await setup(t, { config: BOTH });
  server.queue('zoom_presence', answer(503, ''), 100);
  await advance(30 * MINUTE, 30 * SECOND);
  const [d] = gladys.fake.discoveredDevices;
  const before = gladys.fake.states.length;
  await gladys.fake.deviceCreated(d);
  await settle();
  const sent = gladys.fake.states.slice(before).map((s) => s.device_feature_external_id);
  assert.ok(sent.includes(feature('microsoft_in_meeting')));
  assert.ok(!sent.includes(feature('zoom_in_meeting')), 'Zoom is stale');
  assert.ok(!sent.includes(feature('in_meeting')), 'Teams 0 + Zoom unknown = unknown');
});

test('a very short token: never two renewals within a minute', async (t) => {
  const server = fakeServices();
  server.ms.expiresIn = 20;
  const { advance } = await setup(t, { server });
  await advance(10 * MINUTE, SECOND);
  const renewals = server.of('ms_token').filter((r) => r.form.grant_type === 'refresh_token');
  assert.ok(renewals.length >= 5 && renewals.length <= 11, `${renewals.length} renewals`);
});

test('the wall clock jumping forward never makes the Test button read sooner than the interval', async (t) => {
  const clock = testClock({ start: START });
  let jump = 0;
  const shifted = { ...clock, now: () => clock.now() + jump };
  const { gladys, server } = await setup(t, { clock: shifted });
  const count = server.of('ms_presence').length;
  await clock.advance(15 * SECOND);
  jump = 2 * HOUR;
  const ack = await gladys.fake.action(ACTIONS.TEST_CONNECTION);
  await settle();
  assert.equal(server.of('ms_presence').length, count, JSON.stringify(ack.data));
  assert.match(JSON.stringify(ack.data), /Next read in 4\d s/);
});

test('Connect clicked twice: the return address is shown once', async (t) => {
  const { gladys } = await setup(t, { signIn: [] });
  await gladys.fake.oauthAuthorizeUrl(OAUTH_FIELDS.microsoft, GLADYS_REDIRECT);
  await gladys.fake.oauthAuthorizeUrl(OAUTH_FIELDS.microsoft, GLADYS_REDIRECT);
  await settle();
  assert.equal(message(gladys).en.split('Redirect URI to register').length, 2);
});

test('a wall clock set back never stretches a pause beyond its kind', async (t) => {
  const clock = testClock({ start: START });
  let jump = 0;
  const shifted = { ...clock, now: () => clock.now() - jump };
  const { server, advance } = await setup(t, { clock: shifted });
  server.queue('ms_presence', answer(500, ''));
  await advance(MINUTE);
  const count = server.of('ms_presence').length;
  jump = 10 * HOUR;
  // A 5xx pause is 30 minutes at most, whatever the wall clock says.
  await advance(40 * MINUTE, MINUTE);
  assert.ok(server.of('ms_presence').length > count);
});

test('the guards of a read hold on every path: a tick during a pause, two renewals within a minute', async (t) => {
  const { app, server, advance } = await setup(t);
  server.queue('ms_presence', answer(429, {}, { 'retry-after': '600' }));
  await advance(MINUTE);
  const account = app.accounts.microsoft;
  await advance(2 * MINUTE, MINUTE);
  const count = server.requests.length;
  await account.tick();
  assert.equal(server.requests.length, count, 'no read during a pause');
  await advance(9 * MINUTE, MINUTE);
  const afterPause = server.requests.length;
  await account.renew();
  await assert.rejects(account.renew(), (err) => err.kind === 'renew_wait');
  assert.equal(server.requests.length, afterPause + 1, 'one renewal');
});

test('a sign-in again with the wall clock ahead never reads sooner than the interval', async (t) => {
  const clock = testClock({ start: START });
  let jump = 0;
  const shifted = { ...clock, now: () => clock.now() + jump };
  const { server, authorize } = await setup(t, { clock: shifted });
  const count = server.of('ms_presence').length;
  await clock.advance(10 * SECOND);
  jump = 2 * HOUR;
  await authorize('microsoft');
  await settle();
  assert.equal(server.of('ms_presence').length, count);
});

test('401 on a token renewed for this very read: a stop at once, no second renewal', async (t) => {
  const server = fakeServices();
  server.ms.refresh = 'ms-refresh-0';
  const store = memoryStore({
    microsoft: {
      session: {
        client: MS_CLIENT,
        accessToken: 'ms-access-0',
        refreshToken: 'ms-refresh-0',
        expiresAt: START - HOUR,
        obtainedAt: START - 2 * HOUR,
      },
    },
  });
  server.queue('ms_presence', answer(401, {}), 2);
  await setup(t, { server, store, signIn: [] });
  assert.equal(store.data.microsoft.block?.kind, 'unauthorized');
  assert.equal(server.of('ms_token').length, 1, 'the renewal of the read only');
  assert.equal(server.of('ms_presence').length, 1);
});

test('nothing is ever written while /data could not be read', async (t) => {
  const store = memoryStore();
  store.fail = true;
  const { app } = await setup(t, { store, signIn: [], config: MS });
  await app.saveMemory();
  assert.equal(store.saves, 0);
});
