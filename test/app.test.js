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
import { ACTIONS, LIMITS, OAUTH_FIELDS, TIMINGS } from '../src/constants.js';
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
  // Not awaited: a test that failed with a request held never hangs here.
  t.after(() => {
    app.shutdown().catch(() => {});
  });
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

test('Teams: authorization, one read, device and states; tokens in Gladys only, no secret logged', async (t) => {
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

  // The tokens are kept in the configuration of Gladys, never in /data.
  const kept = gladys.fake.config.microsoft_session;
  assert.equal(kept.client, MS_CLIENT);
  assert.equal(kept.accessToken, server.ms.access);
  assert.equal(kept.refreshToken, server.ms.refresh);
  for (const token of [server.ms.access, server.ms.refresh]) {
    assert.ok(!JSON.stringify(store.data).includes(token), 'no token in /data');
  }
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
  const { gladys, server, advance } = await setup(t);
  const first = server.ms.refresh;
  await advance(55 * MINUTE, 30 * SECOND);
  const renewals = server.of('ms_token').filter((r) => r.form.grant_type === 'refresh_token');
  assert.equal(renewals.length, 1);
  assert.equal(renewals[0].form.refresh_token, first);
  assert.match(renewals[0].form.scope, /offline_access/);
  assert.equal(gladys.fake.config.microsoft_session.refreshToken, server.ms.refresh);
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
  assert.equal(gladys.fake.config.microsoft_session, '');
  assert.equal(store.data.microsoft.block.kind, 'expired');
  const count = server.requests.length;
  await advance(3 * HOUR, 10 * MINUTE);
  assert.equal(server.requests.length, count, 'never retried by itself');
  // A restart on the same /data: still stopped, nothing sent.
  app.stop();
  const again = await setup(t, { store, server, signIn: [], config: gladys.fake.config });
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
  assert.ok(gladys.fake.config.microsoft_session.refreshToken, 'the session is kept');
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
  const again = await setup(t, { store, server, clock, signIn: [], config: gladys.fake.config });
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
  const { app, gladys, store, server, clock } = await setup(t);
  const count = server.of('ms_presence').length;
  await clock.advance(20 * SECOND);
  app.stop();
  const again = await setup(t, { store, server, clock, signIn: [], config: gladys.fake.config });
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
  const again = await setup(t, { store, server, signIn: [], config: first.gladys.fake.config });
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
  const { app, gladys, server, store } = await setup(t, { config: BOTH });
  await gladys.fake.configUpdated({
    ...BOTH,
    microsoft_client_id: 'aaaaaaaa-2222-3333-4444-555555555555',
  });
  await settle();
  assert.match(message(gladys).en, /Microsoft Teams: not signed in.*Zoom: connected/);
  assert.ok(app.accounts.zoom.session.refreshToken);
  assert.ok(!JSON.stringify(store.data).includes('refresh'));
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
    config: first.gladys.fake.config,
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

/** A transport whose requests to `route` wait for `release()`. */
function holdable(server) {
  const base = server.transport;
  const held = { route: null, waiting: [] };
  server.transport = async (url, options) => {
    const route = /token/.test(url) ? 'token' : 'presence';
    if (held.route === route) {
      await new Promise((resolve) => held.waiting.push(resolve));
    }
    return base(url, options);
  };
  held.release = () => {
    held.route = null;
    held.waiting.splice(0).forEach((resolve) => resolve());
  };
  return held;
}

test('a stopped service stays in "Teams or Zoom": an expired Teams never lets it fall to 0', async (t) => {
  const { gladys, server, advance } = await setup(t, { config: BOTH });
  server.zoomPresence = { status: 'In_A_Zoom_Meeting' };
  await advance(MINUTE);
  assert.equal(last(gladys, 'in_meeting'), 1);
  server.queue('ms_token', answer(400, { error: 'invalid_grant' }));
  await advance(55 * MINUTE, 30 * SECOND);
  assert.match(message(gladys).en, /Microsoft Teams: the session ended/);
  server.zoomPresence = { status: 'Available' };
  await advance(10 * MINUTE, 30 * SECOND);
  assert.equal(last(gladys, 'zoom_in_meeting'), 0);
  assert.ok(!statesOf(gladys, 'in_meeting').slice(1).includes(0), 'Teams unknown: never 0');
  // Keys emptied: Teams no longer counts.
  await gladys.fake.configUpdated({ ...ZOOM });
  await settle();
  assert.equal(last(gladys, 'in_meeting'), 0);
});

for (const [name, route, act] of [
  ['an exchange of code', 'token', 'connect'],
  ['a renewal', 'token', 'renew'],
  ['a read', 'presence', 'read'],
]) {
  test(`shutdown during ${name}: no timer nor request afterwards`, async (t) => {
    const server = fakeServices();
    const held = holdable(server);
    const clock = testClock({ start: START });
    const ctx = await setup(t, { server, clock, signIn: act === 'connect' ? [] : undefined });
    let pending;
    if (act === 'connect') {
      const back = await ctx.connect('microsoft');
      held.route = route;
      pending = ctx.gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, back);
    } else {
      held.route = route;
      await ctx.advance(act === 'renew' ? 51 * MINUTE : MINUTE, 30 * SECOND);
    }
    await settle();
    assert.ok(held.waiting.length >= 1, 'a request is held');
    ctx.app.shutdown();
    // The held request itself still goes out; nothing after it.
    const count = server.requests.length + held.waiting.length;
    held.release();
    await pending;
    await settle();
    await clock.advance(5 * HOUR);
    await settle();
    assert.equal(server.requests.length, count, 'no request after shutdown');
    assert.equal(clock.pending, 0, 'no timer after shutdown');
  });
}

test('tokens of an older /data move to the configuration of Gladys, then leave /data', async (t) => {
  const legacy = {
    client: MS_CLIENT,
    accessToken: 'ms-access-7',
    refreshToken: 'ms-refresh-7',
    expiresAt: START + HOUR,
    obtainedAt: START,
  };
  const server = fakeServices();
  server.ms.access = 'ms-access-7';
  server.ms.refresh = 'ms-refresh-7';
  const store = memoryStore({ microsoft: { session: legacy } });
  const { gladys } = await setup(t, { server, store, signIn: [] });
  assert.equal(gladys.fake.config.microsoft_session.refreshToken, 'ms-refresh-7');
  assert.ok(!JSON.stringify(store.data).includes('ms-refresh-7'));
  assert.equal(gladys.fake.connectionStatus.connected, true);
});

test('Gladys refusing to keep the tokens: kept in /data still, saved again later', async (t) => {
  const legacy = {
    client: MS_CLIENT,
    accessToken: 'ms-access-7',
    refreshToken: 'ms-refresh-7',
    expiresAt: START + HOUR,
    obtainedAt: START,
  };
  const server = fakeServices();
  server.ms.access = 'ms-access-7';
  server.ms.refresh = 'ms-refresh-7';
  const store = memoryStore({ microsoft: { session: legacy } });
  const gladysFail = { on: true };
  const clock = testClock({ start: START });
  const gladys = createFakeGladys({ config: MS });
  const setConfig = gladys.setConfig.bind(gladys);
  gladys.setConfig = async (values) => {
    if (gladysFail.on) {
      throw Object.assign(new Error('503'), { status: 503 });
    }
    return setConfig(values);
  };
  const app = new MeetingApp({
    gladys,
    logger: recorder([]),
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
  assert.ok(JSON.stringify(store.data).includes('ms-refresh-7'), 'not lost');
  gladysFail.on = false;
  await clock.advance(2 * MINUTE);
  await settle();
  assert.equal(gladys.fake.config.microsoft_session.refreshToken, 'ms-refresh-7');
});

test('Retry-After: huge capped at 6 h, absent 5 min', async (t) => {
  const { server, advance } = await setup(t);
  server.queue('ms_presence', answer(429, {}, { 'retry-after': '999999999' }));
  await advance(MINUTE);
  const count = server.of('ms_presence').length;
  await advance(5 * HOUR, 30 * MINUTE);
  assert.equal(server.of('ms_presence').length, count);
  await advance(90 * MINUTE, 5 * MINUTE);
  assert.ok(server.of('ms_presence').length > count, 'read again after 6 h');
  server.queue('ms_presence', answer(429, {}));
  await advance(2 * MINUTE, MINUTE);
  const after = server.of('ms_presence').length;
  await advance(3 * MINUTE, MINUTE);
  assert.equal(server.of('ms_presence').length, after);
  await advance(3 * MINUTE, MINUTE);
  assert.ok(server.of('ms_presence').length > after, 'read again after 5 min');
});

test('the requests count in the daily budget; reached, the pause lasts until the next UTC day', async (t) => {
  const { app, gladys, store, server, advance } = await setup(t);
  const before = store.data.microsoft.requests.count;
  await advance(3 * MINUTE, MINUTE);
  assert.equal(store.data.microsoft.requests.count, before + 3);
  app.memory.microsoft.requests.count = LIMITS.dailyRequests;
  const count = server.requests.length;
  await advance(MINUTE, MINUTE);
  assert.match(message(gladys).en, /daily number of requests reached: paused for 1[45] h/);
  // 09:04 UTC: the next UTC day is in about 15 hours.
  await advance(14 * HOUR, 30 * MINUTE);
  assert.equal(server.requests.length, count);
  await advance(2 * HOUR, 10 * MINUTE);
  assert.ok(server.requests.length > count);
});

test('server_error or temporarily_unavailable of a token endpoint is never a refusal', async (t) => {
  const { gladys, store, server, advance } = await setup(t);
  server.queue('ms_token', answer(400, { error: 'temporarily_unavailable' }));
  await advance(55 * MINUTE, 30 * SECOND);
  assert.equal(store.data.microsoft.block, null);
  assert.equal(gladys.fake.connectionStatus.connected, true, 'renewed later');
});

test('during a 429 pause, Connect and the code exchange send nothing', async (t) => {
  const { gladys, server, advance, connect } = await setup(t);
  const back = await connect('microsoft');
  server.queue('ms_presence', answer(429, {}, { 'retry-after': '3600' }));
  await advance(MINUTE);
  const ack = await gladys.fake.oauthAuthorizeUrl(OAUTH_FIELDS.microsoft, GLADYS_REDIRECT);
  assert.equal(ack.success, false);
  const count = server.requests.length;
  const cb = await gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, back);
  assert.equal(cb.success, false);
  assert.equal(server.requests.length, count);
});

test('a code too long, an authorization older than 15 minutes: refused, nothing sent', async (t) => {
  const { gladys, server, connect, advance } = await setup(t, { signIn: [] });
  const back = await connect('microsoft');
  const long = await gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, {
    ...back,
    code: 'x'.repeat(LIMITS.codeLength + 1),
  });
  assert.equal(long.success, false);
  await advance(16 * MINUTE, MINUTE);
  const old = await gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, back);
  assert.equal(old.success, false);
  assert.equal(server.requests.length, 0);
});

test('a stop of other keys never blocks new ones', async (t) => {
  const store = memoryStore({
    microsoft: { block: { kind: 'expired', client: '0123456789abcdef' } },
  });
  const { gladys } = await setup(t, { store, signIn: [] });
  assert.match(message(gladys).en, /not signed in/);
});

test('at most 5 sign-ins wait; one renewal at a time', async (t) => {
  const { app, server, gladys, advance } = await setup(t);
  const urls = [];
  for (let i = 0; i < 6; i += 1) {
    urls.push((await gladys.fake.oauthAuthorizeUrl(OAUTH_FIELDS.microsoft, GLADYS_REDIRECT)).data);
  }
  assert.equal(app.signIns.pending.size, 5);
  const first = server.approve(urls[0].authorize_url);
  assert.equal((await gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, first)).success, false);
  await advance(2 * MINUTE, MINUTE);
  const before = server.of('ms_token').length;
  const account = app.accounts.microsoft;
  await Promise.all([account.renew(), account.renew()]);
  assert.equal(server.of('ms_token').length, before + 1);
});

test('stopped while /data is written before a read: the read is not sent', async (t) => {
  const { app, server, store, advance } = await setup(t);
  let release;
  const save = store.save.bind(store);
  store.save = async (v) => {
    if (release === null) {
      await new Promise((resolve) => (release = resolve));
    }
    return save(v);
  };
  release = null;
  const count = server.of('ms_presence').length;
  const tick = advance(MINUTE, MINUTE);
  await settle();
  app.stop();
  release?.();
  await tick;
  await settle();
  assert.equal(server.of('ms_presence').length, count);
});

test('stop (not shutdown) during an exchange: its answer schedules nothing; after shutdown, no request at all', async (t) => {
  const server = fakeServices();
  const held = holdable(server);
  const clock = testClock({ start: START });
  const { app, connect, gladys } = await setup(t, { server, clock, signIn: [] });
  const back = await connect('microsoft');
  held.route = 'token';
  const pending = gladys.fake.oauthCallback(OAUTH_FIELDS.microsoft, back);
  await settle();
  app.stop();
  held.release();
  await pending;
  await settle();
  const count = server.requests.length;
  await clock.advance(HOUR);
  await settle();
  assert.equal(server.requests.length, count);
  assert.equal(clock.pending, 0);
  await app.shutdown();
  // Even a direct call of the transport of an account sends nothing any more.
  await assert.rejects(
    app.accounts.microsoft.transport('https://graph.microsoft.com/v1.0/me/presence', {}),
    (err) => err.kind === 'stale',
  );
  assert.equal(server.requests.length, count);
});

// --- Tokens kept by Gladys: never a renewal it cannot keep, never one lost ---

/**
 * setConfig of the fake Gladys made refusable: `fail(values)` true refuses
 * (503); `hold(values)` true holds the write until `release()`. Counts the
 * tries; past `limit`, a write never answers (a loop stops, the test fails fast).
 */
function gladysWrites(gladys, limit = 100) {
  const base = gladys.setConfig.bind(gladys);
  const ctl = { calls: 0, fail: () => false, hold: () => false, waiting: [] };
  gladys.setConfig = async (values) => {
    ctl.calls += 1;
    if (ctl.calls > limit) {
      return new Promise(() => {});
    }
    if (ctl.hold(values)) {
      await new Promise((resolve) => ctl.waiting.push(resolve));
    }
    if (ctl.fail(values)) {
      throw Object.assign(new Error('503'), { status: 503 });
    }
    return base(values);
  };
  ctl.release = () => ctl.waiting.splice(0).forEach((resolve) => resolve());
  return ctl;
}

const refreshIn = (gladys) => gladys.fake.config.microsoft_session?.refreshToken;
const newToken = (server) => (values) =>
  values.microsoft_session !== undefined &&
  values.microsoft_session?.refreshToken === server.ms.refresh &&
  server.ms.issued > 1;

/** A promise that rejects after `ms` of real time: a test fails, never hangs. */
const deadline = (promise, ms = 3000) =>
  Promise.race([
    promise,
    new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('still waiting (real deadline)')), ms);
      promise.finally?.(() => clearTimeout(timer));
    }),
  ]);

/** The container restarts: the same /data, the session Gladys kept, the same services. */
async function restart(t, ctx, logs = []) {
  await deadline(ctx.app.shutdown());
  return setup(t, {
    config: { ...ctx.gladys.fake.config },
    store: ctx.store,
    server: ctx.server,
    clock: ctx.clock,
    signIn: [],
    logs,
  });
}

test('Gladys refusing writes: no renewal is sent, a try a minute, renewed once Gladys accepts; a restart meanwhile keeps the session', async (t) => {
  const logs = [];
  const ctx = await setup(t, { logs });
  const { gladys, server, advance } = ctx;
  const writes = gladysWrites(gladys);
  await advance(45 * MINUTE, MINUTE);
  const tokens = server.of('ms_token').length;
  writes.fail = () => true;
  const calls = writes.calls;
  const lines = logs.length;
  // The renewal window (50 min) passes and the token ends (60 min).
  await advance(17 * MINUTE, 10 * SECOND);
  assert.equal(server.of('ms_token').length, tokens, 'no renewal Gladys could not keep');
  const tries = writes.calls - calls;
  assert.ok(tries >= 5 && tries <= 15, `bounded tries: ${tries}`);
  assert.ok(logs.length - lines <= 50, `bounded logs: ${logs.length - lines}`);
  assert.match(message(gladys).en, /Gladys cannot keep the tokens/);
  assert.equal(refreshIn(gladys), 'ms-refresh-1');
  assert.ok(!JSON.stringify(logs).includes('ms-refresh'), 'no token in the logs');
  assert.ok(
    !JSON.stringify(gladys.fake.connectionStatuses).includes('ms-'),
    'no token in a status',
  );
  assert.ok(!JSON.stringify(ctx.store.data).includes('ms-'), 'no token in /data');
  // A restart now: the renewal token Gladys holds still serves.
  const next = await restart(t, ctx);
  await next.advance(3 * MINUTE, 10 * SECOND);
  assert.equal(server.of('ms_token').length, tokens + 1, 'renewed after the restart');
  assert.equal(refreshIn(next.gladys), server.ms.refresh);
  assert.equal(next.store.data.microsoft.block, null);
  assert.equal(next.gladys.fake.connectionStatus.connected, true);
});

test('Gladys accepting again: the renewal follows within a minute', async (t) => {
  const { gladys, server, advance } = await setup(t);
  const writes = gladysWrites(gladys);
  await advance(45 * MINUTE, MINUTE);
  const tokens = server.of('ms_token').length;
  writes.fail = () => true;
  await advance(7 * MINUTE, 10 * SECOND);
  assert.equal(server.of('ms_token').length, tokens);
  writes.fail = () => false;
  await advance(70 * SECOND, 10 * SECOND);
  assert.equal(server.of('ms_token').length, tokens + 1);
  assert.equal(refreshIn(gladys), server.ms.refresh);
});

test('Gladys not connected: no renewal and no write, a try a minute; at the reconnection, the renewal follows', async (t) => {
  const logs = [];
  const { gladys, server, advance } = await setup(t, { logs });
  const writes = gladysWrites(gladys);
  await advance(49 * MINUTE, MINUTE);
  const tokens = server.of('ms_token').length;
  await gladys.disconnect();
  const calls = writes.calls;
  // Checked early: a renewal loop fails here within seconds.
  await advance(2 * MINUTE, 30 * SECOND);
  const tries = logs.filter((l) => /token renewal failed/.test(l)).length;
  assert.ok(tries >= 1 && tries <= 3, `a try a minute, never a loop: ${tries}`);
  await advance(8 * MINUTE, 10 * SECOND);
  assert.equal(server.of('ms_token').length, tokens, 'no renewal while Gladys is away');
  assert.equal(writes.calls, calls, 'no write while Gladys is away');
  await gladys.connect();
  await advance(70 * SECOND, 10 * SECOND);
  assert.equal(server.of('ms_token').length, tokens + 1);
  assert.equal(refreshIn(gladys), server.ms.refresh);
});

test('Gladys refusing the new token: saved again by a timer within seconds (not at the next read); a restart after keeps it', async (t) => {
  const ctx = await setup(t, { config: { ...MS, poll_seconds: 900 } });
  const { gladys, server, advance } = ctx;
  const writes = gladysWrites(gladys);
  writes.fail = newToken(server);
  await advance(51 * MINUTE, 10 * SECOND);
  assert.equal(server.ms.refresh, 'ms-refresh-2', 'renewed (the witness write was accepted)');
  assert.equal(refreshIn(gladys), 'ms-refresh-1', 'the new token was refused');
  assert.equal(ctx.app.accounts.microsoft.sessionPending, true);
  // Refused for a while: tried again every 10 s to a minute, never another renewal.
  await advance(3 * MINUTE, 5 * SECOND);
  assert.equal(server.of('ms_token').length, 2);
  writes.fail = () => false;
  const reads = server.of('ms_presence').length;
  await advance(MINUTE, 5 * SECOND);
  assert.equal(refreshIn(gladys), 'ms-refresh-2');
  assert.equal(server.of('ms_presence').length, reads, 'saved by the timer, not by a read');
  assert.ok(!JSON.stringify(ctx.store.data).includes('ms-refresh'), 'no token in /data');
  const next = await restart(t, ctx);
  await next.advance(20 * MINUTE, MINUTE);
  assert.equal(next.store.data.microsoft.block, null, 'never "session ended"');
  assert.equal(next.gladys.fake.connectionStatus.connected, true);
  assert.equal(next.app.accounts.microsoft.session.refreshToken, server.ms.refresh);
});

test('a new token Gladys refused is saved at once at the reconnection to Gladys', async (t) => {
  const { app, gladys, server, advance } = await setup(t, { config: { ...MS, poll_seconds: 900 } });
  const writes = gladysWrites(gladys);
  writes.fail = newToken(server);
  await advance(51 * MINUTE, 10 * SECOND);
  assert.equal(refreshIn(gladys), 'ms-refresh-1');
  await gladys.disconnect();
  writes.fail = () => false;
  await gladys.connect();
  await settle();
  assert.equal(refreshIn(gladys), 'ms-refresh-2');
  assert.equal(app.accounts.microsoft.sessionPending, false);
});

test('one write at a time: an older write accepted never clears a newer session refused', async (t) => {
  const { app, gladys, clock } = await setup(t);
  const account = app.accounts.microsoft;
  const writes = gladysWrites(gladys);
  const s = account.session;
  const a = { ...s, refreshToken: 'ms-refresh-A' };
  const b = { ...s, refreshToken: 'ms-refresh-B' };
  writes.hold = (v) => v.microsoft_session?.refreshToken === 'ms-refresh-A';
  writes.fail = (v) => v.microsoft_session?.refreshToken === 'ms-refresh-B';
  const first = account.setSession(a);
  await settle();
  const second = account.setSession(b);
  writes.release();
  assert.equal(await first, true);
  assert.equal(await second, false);
  assert.equal(refreshIn(gladys), 'ms-refresh-A');
  assert.equal(account.sessionPending, true, 'the newer session is still to save');
  writes.fail = () => false;
  await clock.advance(10 * SECOND);
  await settle();
  assert.equal(refreshIn(gladys), 'ms-refresh-B');
  assert.equal(account.sessionPending, false);
});

test('a write failing after stop() arms no timer; the next start tries it again', async (t) => {
  const { app, gladys, clock } = await setup(t);
  const account = app.accounts.microsoft;
  const writes = gladysWrites(gladys);
  writes.hold = () => true;
  writes.fail = () => true;
  const pending = account.setSession({ ...account.session, refreshToken: 'ms-refresh-C' });
  await settle();
  app.stop();
  writes.hold = () => false;
  writes.release();
  assert.equal(await pending, false);
  await settle();
  assert.equal(clock.pending, 0, 'no timer after stop()');
  writes.fail = () => false;
  await app.start(app.rawConfig);
  await clock.advance(10 * SECOND);
  await settle();
  assert.equal(refreshIn(gladys), 'ms-refresh-C');
});

test('shutdown (the SDK cleanup) during a renewal: waited for, the new token kept by Gladys, a restart keeps the session', async (t) => {
  const server = fakeServices();
  const held = holdable(server);
  const ctx = await setup(t, { server });
  held.route = 'token';
  await ctx.advance(51 * MINUTE, 30 * SECOND);
  await settle();
  assert.equal(held.waiting.length, 1, 'a renewal is in flight');
  const done = ctx.gladys.fake.shutdown();
  await settle();
  held.release();
  await deadline(done);
  assert.equal(server.ms.refresh, 'ms-refresh-2');
  assert.equal(refreshIn(ctx.gladys), 'ms-refresh-2', 'the new renewal token is in Gladys');
  assert.equal(ctx.clock.pending, 0, 'no timer after shutdown');
  const next = await restart(t, ctx);
  await next.advance(20 * MINUTE, MINUTE);
  assert.equal(next.store.data.microsoft.block, null, 'never "session ended"');
  assert.equal(next.gladys.fake.connectionStatus.connected, true);
});

test('shutdown waits 8 s at most for a renewal (Docker kills after 10 s), then nothing runs', async (t) => {
  assert.ok(TIMINGS.shutdownWait <= 8 * SECOND);
  const server = fakeServices();
  const held = holdable(server);
  const ctx = await setup(t, { server });
  held.route = 'token';
  await ctx.advance(51 * MINUTE, 30 * SECOND);
  await settle();
  assert.equal(held.waiting.length, 1);
  let finished = false;
  const done = ctx.app.shutdown().then(() => (finished = true));
  await settle();
  await ctx.clock.advance(7.9 * SECOND);
  await settle();
  assert.equal(finished, false, 'still waiting for the renewal');
  await ctx.clock.advance(0.2 * SECOND);
  await deadline(done);
  assert.equal(ctx.clock.pending, 0);
  // The answer arriving later is not written (the container is gone).
  const kept = ctx.gladys.fake.config.microsoft_session;
  held.release();
  await settle();
  assert.equal(ctx.gladys.fake.config.microsoft_session, kept);
  assert.equal(ctx.clock.pending, 0);
});

test('no renewal leaves while the container stops: a read in flight refused (401) sends no token request nor write', async (t) => {
  const server = fakeServices();
  const held = holdable(server);
  const ctx = await setup(t, { server });
  const writes = gladysWrites(ctx.gladys);
  await ctx.advance(10 * MINUTE, MINUTE);
  held.route = 'presence';
  server.queue('ms_presence', answer(401, { error: { code: 'InvalidAuthenticationToken' } }));
  await ctx.advance(MINUTE, MINUTE);
  await settle();
  assert.equal(held.waiting.length, 1, 'a read is in flight');
  const tokens = server.of('ms_token').length;
  const calls = writes.calls;
  const done = ctx.app.shutdown();
  held.release();
  await deadline(done);
  await settle();
  assert.equal(server.of('ms_token').length, tokens, 'no renewal during the stop');
  assert.equal(writes.calls, calls, 'no write during the stop');
});

test('the container stopping during the witness write: the renewal is not sent', async (t) => {
  const { app, gladys, server, advance } = await setup(t);
  const writes = gladysWrites(gladys);
  await advance(45 * MINUTE, MINUTE);
  const tokens = server.of('ms_token').length;
  writes.hold = () => true;
  await advance(6 * MINUTE, 30 * SECOND);
  assert.equal(writes.waiting.length, 1, 'the witness write is in flight');
  const done = app.shutdown();
  writes.hold = () => false;
  writes.release();
  await deadline(done);
  await settle();
  assert.equal(server.of('ms_token').length, tokens);
  assert.equal(refreshIn(gladys), 'ms-refresh-1');
});

test('one write at a time: an older write in flight never lands after a newer one', async (t) => {
  const { app, gladys } = await setup(t);
  const account = app.accounts.microsoft;
  const writes = gladysWrites(gladys);
  const s = account.session;
  writes.hold = (v) => v.microsoft_session?.refreshToken === 'ms-refresh-A';
  const first = account.setSession({ ...s, refreshToken: 'ms-refresh-A' });
  await settle();
  const second = account.setSession({ ...s, refreshToken: 'ms-refresh-B' });
  await settle();
  writes.release();
  await deadline(Promise.all([first, second]));
  assert.equal(refreshIn(gladys), 'ms-refresh-B');
  assert.equal(account.sessionPending, false);
});

test('a new sign-in during the witness write: the older session is not renewed', async (t) => {
  const ctx = await setup(t);
  const { gladys, server, advance } = ctx;
  const writes = gladysWrites(gladys);
  await advance(45 * MINUTE, MINUTE);
  writes.hold = (v) => v.microsoft_session?.refreshToken === 'ms-refresh-1';
  await advance(6 * MINUTE, 30 * SECOND);
  assert.equal(writes.waiting.length, 1, 'the witness write is in flight');
  const tokens = server.of('ms_token').length;
  const signIn = ctx.authorize('microsoft');
  await settle();
  writes.hold = () => false;
  writes.release();
  await deadline(signIn);
  await settle();
  assert.equal(server.of('ms_token').length, tokens + 1, 'the code exchange only');
  assert.equal(refreshIn(gladys), server.ms.refresh);
});

test('no read starts while the container stops', async (t) => {
  const server = fakeServices();
  const held = holdable(server);
  const ctx = await setup(t, { server });
  held.route = 'token';
  await ctx.advance(51 * MINUTE, 30 * SECOND);
  await settle();
  assert.equal(held.waiting.length, 1, 'a renewal keeps the stop waiting');
  const done = ctx.app.shutdown();
  const account = ctx.app.accounts.microsoft;
  account.lastReadMono = -Infinity;
  const reads = server.of('ms_presence').length;
  await account.tick();
  assert.equal(server.of('ms_presence').length, reads);
  held.release();
  await deadline(done);
});
