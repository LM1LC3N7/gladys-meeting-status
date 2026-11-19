// -----------------------------------------------------------------------------
// Mutants of the guards of gladys-meeting-status. Each one breaks one guard;
// the whole suite must fail, never hang. Run one at a time in a private copy
// of the repository (never the working tree), under a RAM cap and a deadline
// that kills the whole process group:
//
//   node scripts/mutants.mjs [names...]
//   MUTANTS_REPO=<another clone> node scripts/mutants.mjs   (default: this repository)
//
// The unmodified suite runs first in the copy: red, the run stops ("baseline
// FAILED", exit 1) and no mutant is judged.
//
// Needs `memcap`, `setsid` and `timeout` (the dev VM). Not in the image.
// -----------------------------------------------------------------------------
import { cpSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { dirname, join, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = resolve(
  process.env.MUTANTS_REPO ?? join(dirname(fileURLToPath(import.meta.url)), '..'),
);
const only = process.argv.slice(2);

const A = 'src/account.js';
const P = 'src/app.js';
const O = 'src/oauth.js';
const R = 'src/presence.js';
const H = 'src/http.js';
const D = 'src/devices.js';

/** [name, [file, from, to], ...]: each `from` must occur exactly once. */
const MUTANTS = [
  ['mono-interval-guard', [A, 'if (sinceLast < this.interval) {', 'if (false) {']],
  [
    'pause-ignored',
    [
      A,
      '    if (now < this.pausedUntil) {\n      this.scheduleNext();',
      '    if (false) {\n      this.scheduleNext();',
    ],
  ],
  [
    'renew-floor',
    [A, 'if (now < (this.memory.lastRenewAt ?? -Infinity) + TIMINGS.renewFloor) {', 'if (false) {'],
  ],
  [
    'renew-refused-not-blocked',
    [
      A,
      "      if (err.kind === 'refused') {\n        // Never tried again",
      '      if (false) {\n        // Never tried again',
    ],
  ],
  [
    'refresh-not-rotated',
    [A, 'refreshToken: token.refreshToken ?? s.refreshToken,', 'refreshToken: s.refreshToken,'],
  ],
  [
    'fresh-token-401-renews',
    [A, 'if ((this.session?.obtainedAt ?? 0) >= this.cycleStart) {', 'if (false) {'],
  ],
  ['cycle-not-persisted', [A, '    this.memory.lastCycleAt = now;\n', '']],
  [
    'budget-off-by-one',
    [
      A,
      'if (memory.requests.count >= LIMITS.dailyRequests) {',
      'if (memory.requests.count > LIMITS.dailyRequests) {',
    ],
  ],
  ['stale-reading-used', [A, 'if (mono - this.reading.at > limit) {', 'if (false) {']],
  [
    'refused-without-oauth-code',
    [
      O,
      'if ((res.status === 400 || res.status === 401) && code) {',
      'if (res.status === 400 || res.status === 401) {',
    ],
  ],
  [
    'state-reusable',
    [
      O,
      '    this.pending.delete(state);\n    return this.valid(entry)',
      '    return this.valid(entry)',
    ],
  ],
  [
    'state-other-service',
    [P, 'if (!provider || this.signIns.peek(inner)?.provider !== provider) {', 'if (!provider) {'],
  ],
  [
    'pasted-base-unchecked',
    [
      P,
      'if (parsed.base !== null && parsed.base !== baseOf(pending.redirectUri)) {',
      'if (false) {',
    ],
  ],
  [
    'write-over-unreadable',
    [
      P,
      '    if (!this.loaded) {\n      // Never written',
      '    if (false) {\n      // Never written',
    ],
  ],
  [
    'ms-unknown-as-zero',
    [R, 'if (!MS_AVAILABILITY.has(availability) || !MS_ACTIVITY.has(activity)) {', 'if (false) {'],
  ],
  ['zoom-unknown-as-zero', [R, 'if (!ZOOM_STATUS.has(status)) {', 'if (false) {']],
  ['zoom-5411-not-forbidden', [R, 'if (Number(data?.code) === 5411) {', 'if (false) {']],
  ['any-some-zero', [D, 'values.every((v) => v === 0)', 'values.some((v) => v === 0)']],
  ['size-unbounded', [H, 'if (size > maxBytes) {', 'if (false) {']],
  ['retry-after-local-clock', [H, 'Number.isNaN(reference) ? now : reference', 'now']],
  [
    'pause-not-capped',
    [A, '    const now = this.now();\n    this.capPause(now);\n', '    const now = this.now();\n'],
  ],
  [
    'connect-hint-grows',
    [P, 'join(account.stopMessage(), T.signIn', 'join(account.status?.message, T.signIn'],
  ],
  [
    'any-keeps-stopped-services',
    [
      A,
      '(Boolean(this.session) || this.host.memory(this.provider)?.block?.client === client)',
      'Boolean(this.session)',
    ],
  ],
  [
    'stop-keeps-epoch',
    [
      A,
      '    // An exchange or a renewal in flight never schedules anything after it.\n    this.epoch += 1;\n',
      '',
    ],
  ],
  [
    'closed-transport-open',
    [
      A,
      "      if (this.closed) {\n        throw new ServiceError('stale', 'stopped');",
      "      if (false) {\n        throw new ServiceError('stale', 'stopped');",
    ],
  ],
  [
    'tokens-not-saved',
    [A, '    this.sessionPending = true;\n    return this.saveSession();', '    return true;'],
  ],
  ['legacy-tokens-left', [P, '        delete this.memory[p].legacySession;\n', '']],
  ['state-ttl-off', [O, 'return age >= 0 && age <= TIMINGS.signInTtl;', 'return age >= 0;']],
  [
    'retry-after-no-max',
    [
      A,
      '      wait = Math.min(\n        TIMINGS.rateLimitMax,\n        Math.max(TIMINGS.rateLimitMin, err.retryAfter ?? TIMINGS.rateLimitDefault),\n      );',
      '      wait = Math.max(TIMINGS.rateLimitMin, err.retryAfter ?? TIMINGS.rateLimitDefault);',
    ],
  ],
  [
    'retry-after-default-0',
    [A, 'err.retryAfter ?? TIMINGS.rateLimitDefault', 'err.retryAfter ?? 0'],
  ],
  ['count-not-incremented', [A, '    memory.requests.count += 1;\n', '']],
  ['budget-pause-short', [A, 'wait = (dayOf(now) + 1) * DAY - now + MINUTE;', 'wait = MINUTE;']],
  ['server-error-refused', [O, 'if (TRANSIENT_CODES.has(code)) {', 'if (false) {']],
  [
    'hardpaused-exchange',
    [A, '    if (this.hardPaused) {\n      restore();', '    if (false) {\n      restore();'],
  ],
  ['hardpaused-authorize', [P, 'if (account.hardPaused) {', 'if (false) {']],
  ['code-length-unbounded', [P, 'code.length > LIMITS.codeLength', 'false']],
  [
    'block-client-ignored',
    [A, 'if (this.memory.block?.client === client) {', 'if (this.memory.block) {'],
  ],
  ['signin-cap', [O, 'while (this.pending.size >= LIMITS.signIns) {', 'while (false) {']],
  ['single-flight-renew', [A, 'if (!this.renewing) {', 'if (true) {']],
  [
    'generation-after-save',
    [
      A,
      '    await this.host.saveMemory();\n    if (generation !== this.generation) {\n      return;\n    }\n    try {',
      '    await this.host.saveMemory();\n    try {',
    ],
  ],
  // Tokens kept by Gladys: never a renewal it cannot keep, never one lost.
  [
    'renew-without-witness',
    [
      A,
      'if (!this.host.gladysConnected() || !(await this.saveSession({ witness: true }))) {',
      'if (false) {',
    ],
  ],
  ['renew-gladys-away', [A, 'if (!this.host.gladysConnected() || !(await', 'if (!(await']],
  [
    'witness-not-written',
    [A, 'if (!this.sessionPending && !witness) {', 'if (!this.sessionPending) {'],
  ],
  ['gladys-wait-loop', [A, '      this.renewRetryAt = this.now() + TIMINGS.gladysWait;\n', '']],
  [
    'renew-replaced-session',
    [
      A,
      "    if (s !== this.session) {\n      throw new ServiceError('renew_wait', 'a new sign-in",
      "    if (false) {\n      throw new ServiceError('renew_wait', 'a new sign-in",
    ],
  ],
  [
    'save-no-retry-timer',
    [
      A,
      '      if (this.sessionPending && generation === this.generation) {\n        this.scheduleSessionSave();',
      '      if (false) {\n        this.scheduleSessionSave();',
    ],
  ],
  [
    'save-timer-after-stop',
    [
      A,
      'if (this.sessionPending && generation === this.generation) {',
      'if (this.sessionPending) {',
    ],
  ],
  [
    'older-write-clears-pending',
    [
      A,
      '    if (this.stored === value) {\n      this.sessionPending = false;',
      '    if (true) {\n      this.sessionPending = false;',
    ],
  ],
  [
    'writes-not-chained',
    [
      A,
      'const link = this.sessionChain.then(() => this.writeSession(witness, generation));',
      'const link = this.writeSession(witness, generation);',
    ],
  ],
  [
    'start-no-pending-retry',
    [
      A,
      '    if (this.sessionPending) {\n      // Not saved before this start',
      '    if (false) {\n      // Not saved before this start',
    ],
  ],
  [
    'reconnect-no-save',
    [P, '    await Promise.all(PROVIDERS.map((p) => this.accounts[p].saveSession()));\n', ''],
  ],
  [
    'shutdown-no-wait',
    [P, '    await Promise.race([Promise.all(accounts.map((a) => a.settle())), late]);\n', ''],
  ],
  [
    'shutdown-unbounded',
    [
      P,
      'await Promise.race([Promise.all(accounts.map((a) => a.settle())), late]);',
      'await Promise.all(accounts.map((a) => a.settle()));',
    ],
  ],
  [
    'shutdown-closed-early',
    [P, 'account.closing = true;', 'account.closing = true;\n      account.closed = true;'],
  ],
  [
    'shutdown-not-wired',
    [
      'src/handlers.js',
      'gladys.handleShutdown(() => app.shutdown());',
      'gladys.handleShutdown(() => {});',
    ],
  ],
  [
    'renew-while-closing-request',
    [
      A,
      '    if (this.closing) {\n      // The container stops: its token could not be kept',
      '    if (false) {\n      // The container stops: its token could not be kept',
    ],
  ],
  ['read-while-closing', [A, 'this.closed || this.closing) {', 'this.closed) {']],
];

function apply(file, from, to) {
  const text = readFileSync(file, 'utf8');
  const count = text.split(from).length - 1;
  if (count !== 1) {
    throw new Error(`${count} occurrences of ${JSON.stringify(from.slice(0, 60))}`);
  }
  writeFileSync(
    file,
    text.replace(from, () => to),
  );
}

/** The whole suite in `cwd`, its process group killed at the deadline. */
function suite(cwd) {
  const started = Date.now();
  const run = spawnSync(
    'setsid',
    ['memcap', '-m', '3G', 'timeout', '-k', '5', '300', 'node', '--test', '--test-timeout=60000'],
    { cwd, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 },
  );
  const out = `${run.stdout}${run.stderr}`;
  return {
    status: run.status,
    failed: /ℹ fail (\d+)/.exec(out)?.[1],
    hung: run.status === 124 || run.status === 137,
    seconds: Math.round((Date.now() - started) / 1000),
  };
}

// A private copy (node_modules copied, never linked), removed at the end.
// Only `.git` itself is left out (`.github` holds the release script a test imports).
const git = join(repo, '.git');
const work = mkdtempSync(join(tmpdir(), 'meeting-mutants-'));
let survivors = 0;
try {
  const copy = join(work, 'repo');
  cpSync(repo, copy, {
    recursive: true,
    filter: (src) => src !== git && !src.startsWith(`${git}${sep}`),
  });
  const base = suite(copy);
  console.log(`baseline: exit ${base.status}, fail ${base.failed}, ${base.seconds} s`);
  if (base.status !== 0 || base.failed !== '0') {
    console.log('baseline FAILED: the unmodified suite is red, no mutant judged');
    survivors = Infinity;
  }
  for (const [name, ...edits] of survivors === 0 ? MUTANTS : []) {
    if (only.length > 0 && !only.includes(name)) {
      continue;
    }
    const originals = new Map();
    let line;
    try {
      for (const [file, from, to] of edits) {
        const path = join(copy, file);
        if (!originals.has(path)) {
          originals.set(path, readFileSync(path, 'utf8'));
        }
        apply(path, from, to);
      }
      const run = suite(copy);
      // Judged on the exit status: a failed test, a test past its deadline
      // (cancelled, not "fail") or the whole run past its deadline all kill it.
      const killed = run.status !== 0;
      if (!killed) {
        survivors += 1;
      }
      const verdict = run.hung
        ? `KILLED, deadline (${run.status})`
        : killed
          ? 'KILLED'
          : 'SURVIVED';
      line = `${name}: ${verdict} (fail ${run.failed}, ${run.seconds} s)`;
    } catch (err) {
      survivors += 1;
      line = `${name}: ERROR ${err.message}`;
    } finally {
      for (const [path, original] of originals) {
        writeFileSync(path, original);
      }
    }
    console.log(line);
  }
} finally {
  rmSync(work, { recursive: true, force: true });
}
process.exitCode = survivors > 0 ? 1 : 0;
