// The period on the real clock (sped up): one presence read a minute, never
// faster, measured on the decision (Account.run). Everything is stopped in
// t.after, so a failure never leaves a timer behind.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createFakeGladys } from '@gladysassistant/integration-sdk/testing';
import { MeetingApp } from '../src/app.js';
import { measurePeriod, realClock, spyOn } from './kit/index.mjs';
import { MS_CLIENT_ID, MS_SECRET, fakeServices, memoryStore } from './helpers/fakeServices.js';

const silent = { debug() {}, info() {}, warn() {}, error() {} };

test('one presence read a minute on the real clock, never faster', async (t) => {
  const clock = realClock({ speed: 200 });
  const server = fakeServices();
  const store = memoryStore({
    microsoft: {
      session: {
        client: (await import('../src/config.js')).clientKey(
          'microsoft',
          MS_CLIENT_ID,
          'organizations',
        ),
        accessToken: 'ms-access-1',
        refreshToken: 'ms-refresh-1',
        expiresAt: clock.now() + 24 * 3600 * 1000,
        obtainedAt: clock.now(),
      },
    },
  });
  server.ms.access = 'ms-access-1';
  server.ms.refresh = 'ms-refresh-1';
  const gladys = createFakeGladys();
  const app = new MeetingApp({
    gladys,
    logger: silent,
    transport: server.transport,
    now: clock.now,
    mono: clock.monotonic,
    timers: clock.timers,
    store,
  });
  t.after(() => {
    app.stop();
    clock.close();
  });
  const spy = spyOn(app.accounts.microsoft, 'call', clock);
  await app.start({ microsoft_client_id: MS_CLIENT_ID, microsoft_client_secret: MS_SECRET });
  const gaps = await measurePeriod(spy, {
    count: 4,
    period: 60_000,
    max: 60_000 * 3,
    timeoutMs: 20_000,
  });
  assert.equal(gaps.length, 3);
  assert.ok(
    gaps.every((g) => g >= 60_000 * 0.95),
    JSON.stringify(gaps),
  );
});
