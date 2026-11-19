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
      const killed = !run.hung && run.failed !== undefined && run.failed !== '0';
      if (!killed) {
        survivors += 1;
      }
      const verdict = run.hung ? `HUNG(${run.status})` : killed ? 'KILLED' : 'SURVIVED';
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
