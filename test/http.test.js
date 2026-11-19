// -----------------------------------------------------------------------------
// The transport against a real local HTTPS server (the kit's): an answer too
// large is cut before it is read whole, a silent server ends at the timeout,
// a socket error is a network error without the request in it.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import https from 'node:https';
import { ServiceError, httpsTransport, parseRetryAfter } from '../src/http.js';
import { blackhole, certs, httpsServer, reply } from './kit/index.mjs';

const agentFor = (cert) =>
  new https.Agent({ ca: cert.cert, checkServerIdentity: () => undefined, keepAlive: false });

test('an answer larger than the bound is refused before it is read whole', async (t) => {
  const srv = await httpsServer({ routes: { '/p': () => reply(200, '{}') }, bigBody: 200_000 });
  t.after(() => srv.close());
  await assert.rejects(
    httpsTransport(`${srv.url}/p`, { agent: agentFor(certs.device), maxBytes: 64 * 1024 }),
    (err) => err instanceof ServiceError && err.kind === 'size',
  );
});

test('a normal answer: status, headers, body', async (t) => {
  const srv = await httpsServer({
    routes: { '/p': () => reply(200, '{"status":"Available"}') },
  });
  t.after(() => srv.close());
  const res = await httpsTransport(`${srv.url}/p`, { agent: agentFor(certs.device) });
  assert.equal(res.status, 200);
  assert.equal(JSON.parse(res.body).status, 'Available');
});

test('a server that never answers ends at the timeout', async (t) => {
  const hole = await blackhole();
  t.after(() => hole.close());
  await assert.rejects(
    httpsTransport(`https://127.0.0.1:${hole.port}/p`, { timeout: 300 }),
    (err) => err.kind === 'timeout',
  );
});

test('only HTTPS is used', async () => {
  await assert.rejects(httpsTransport('http://127.0.0.1:1/p'), (err) => err.kind === 'network');
});

test('Retry-After: seconds, or a date against the Date header of the answer', () => {
  assert.equal(parseRetryAfter('120', 0), 120_000);
  assert.equal(parseRetryAfter(undefined, 0), undefined);
  const server = Date.UTC(2030, 0, 1);
  const date = new Date(server + 90_000).toUTCString();
  assert.equal(parseRetryAfter(date, 0, new Date(server).toUTCString()), 90_000);
  assert.equal(parseRetryAfter(date, server + 30_000), 60_000, 'without Date: our clock');
  assert.equal(parseRetryAfter('soon', 0), undefined);
});
