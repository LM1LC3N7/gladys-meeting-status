// -----------------------------------------------------------------------------
// What must survive a restart, in /data (the only writable place of the
// container), per account (Microsoft, Zoom). Never a password (typed on the
// service's page only), never the client secret (kept by Gladys), never a
// presence (a restart starts from "unknown"):
//   - the session: the access and renewal tokens with their ends, and the
//     client they belong to (a hash of service, client id and tenant);
//   - a stop that only the user lifts (renewal refused, access refused with
//     a new token), with its client, so a restart never retries it;
//   - when the last read succeeded (a restart before the next read stays
//     connected), cleared by any failure, pause or stop;
//   - the end of a 429, 5xx, 403, daily budget or error pause, capped when
//     read again, and when the last read started (a restart never reads
//     sooner than the interval);
//   - when the token was last renewed (never two renewals within a minute);
//   - the requests sent today (the daily budget).
// The tokens are kept as the services give them, in clear, in a file
// readable by the container user only (0600).
// -----------------------------------------------------------------------------

import { open, readFile, rename } from 'node:fs/promises';
import { DAY, LIMITS, PROVIDERS, TIMINGS } from './constants.js';

export const STATE_FILE = '/data/meeting-status-state.json';

/**
 * The file. `load()` resolves the object, `null` when there is no file yet,
 * and throws when it exists but cannot be read (never taken for "empty").
 */
export function fileStore(path = STATE_FILE) {
  // One write at a time: two saves never share the temporary file.
  let chain = Promise.resolve();
  return {
    async load() {
      // After the writes already asked: never an older file than the last save.
      await chain;
      let text;
      try {
        text = await readFile(path, 'utf8');
      } catch (err) {
        if (err?.code === 'ENOENT') {
          return null;
        }
        throw err;
      }
      return JSON.parse(text);
    },
    save(data) {
      const text = JSON.stringify(data);
      const write = chain.then(async () => {
        const temporary = `${path}.tmp`;
        const file = await open(temporary, 'w', 0o600);
        try {
          await file.chmod(0o600);
          await file.writeFile(text);
          // On disk before it replaces the previous state (a power cut keeps one of the two).
          await file.sync();
        } finally {
          await file.close();
        }
        await rename(temporary, path);
      });
      chain = write.catch(() => {});
      return write;
    },
  };
}

const time = (value) => (typeof value === 'number' && Number.isFinite(value) ? value : undefined);
const str = (value, max) =>
  typeof value === 'string' && value !== '' && value.length <= max ? value : undefined;
const client = (value) =>
  typeof value === 'string' && /^[0-9a-f]{16}$/.test(value) ? value : undefined;

const BLOCKS = new Set(['expired', 'unauthorized']);
export const PAUSES = {
  rate: TIMINGS.rateLimitMax,
  unavailable: TIMINGS.unavailable.at(-1),
  forbidden: TIMINGS.forbidden,
  budget: DAY,
  backoff: TIMINGS.backoffMax,
};

function session(raw) {
  const s = raw?.session;
  const owner = client(s?.client);
  const refreshToken = str(s?.refreshToken, LIMITS.tokenLength);
  if (!owner || !refreshToken) {
    return null;
  }
  return {
    client: owner,
    accessToken: str(s.accessToken, LIMITS.tokenLength) ?? null,
    refreshToken,
    expiresAt: time(s.expiresAt) ?? 0,
    obtainedAt: time(s.obtainedAt) ?? 0,
  };
}

/** The UTC day of a time: the daily budget counts from 00:00 UTC. */
export const dayOf = (now) => Math.floor(now / DAY);

/** The persisted state of one account, repaired; a pause end is capped from `now`. */
export function normalizeAccount(raw, now) {
  const block = raw?.block;
  const pause = raw?.pause;
  const until = time(pause?.until);
  const lastRenewAt = time(raw?.lastRenewAt);
  const lastCycleAt = time(raw?.lastCycleAt);
  const lastSuccessAt = time(raw?.lastSuccessAt);
  const requests = raw?.requests;
  return {
    session: session(raw),
    block:
      BLOCKS.has(block?.kind) && client(block?.client)
        ? { kind: block.kind, client: block.client }
        : null,
    pause:
      pause && Object.hasOwn(PAUSES, pause.kind) && until !== undefined && until > now
        ? { kind: pause.kind, until: Math.min(until, now + PAUSES[pause.kind]) }
        : null,
    // A time ahead of the clock counts as now (never a renewal blocked for long).
    lastRenewAt: lastRenewAt === undefined ? null : Math.min(lastRenewAt, now),
    // The start of the last read: the first read after a restart waits the interval.
    lastCycleAt: lastCycleAt === undefined ? null : Math.min(lastCycleAt, now),
    // The start of the last read that succeeded, cleared by any failure, pause or stop.
    lastSuccessAt: lastSuccessAt === undefined ? null : Math.min(lastSuccessAt, now),
    // The requests of today only; another day starts from 0.
    requests:
      requests && requests.day === dayOf(now) && Number.isSafeInteger(requests.count)
        ? { day: requests.day, count: Math.max(0, requests.count) }
        : { day: dayOf(now), count: 0 },
  };
}

/** The whole file: one entry per account. */
export function normalizeMemory(raw, now) {
  return Object.fromEntries(
    PROVIDERS.map((provider) => [provider, normalizeAccount(raw?.[provider], now)]),
  );
}
