// -----------------------------------------------------------------------------
// One account (Microsoft or Zoom): its session (OAuth2 of the app the user
// registered, started by a click), the token renewed by a timer before its
// end, one read of the presence per interval (a minute at least), its
// pauses and stops, its status line. The model is the one of gladys-oura
// (approved), one account per service.
//
// Never by itself: a sign-in, a retry of a refused renewal or access, a read
// sooner than the interval (after an error, a test click, a reconnection to
// Gladys, a configuration saved unchanged or a restart: the start of the
// last read is kept in /data, and a monotonic clock guards the run), two
// renewals or two reads at once, more than LIMITS.dailyRequests requests a
// day. Pauses (429, 5xx, 403, daily budget, errors) and stops are kept in
// /data.
// -----------------------------------------------------------------------------

import { LIMITS, MINUTE, SECOND, TIMINGS, DAY } from './constants.js';
import { ServiceError, errorReason, networkError } from './http.js';
import { T, age, join } from './i18n.js';
import { authorizeUrl, tokenRequest } from './oauth.js';
import { inMeetingOf, readPresence, textOf } from './presence.js';
import { PAUSES, dayOf } from './store.js';

/** The stops kept in /data: only the user lifts them. */
const PERSISTED = new Set(['expired', 'unauthorized']);
const STOP_MESSAGES = {
  no_client: T.incomplete,
  signed_out: T.signedOut,
  expired: T.expired,
  unauthorized: T.unauthorized,
};
/** The pauses that hold every path (the sign-in, the renewal, the reads). */
const HARD = new Set(['rate', 'unavailable', 'budget']);
/** The errors that pause every path. */
const PAUSING = new Set(['rate_limited', 'unavailable', 'budget']);
/** Errors after which the account counts as unreachable (badge). */
const OFFLINE = new Set(['network', 'timeout', 'unavailable', 'bad_response', 'size', 'http']);
/** An access token closer than this to its end is not used (renewed first). */
const TOKEN_SLACK = 10 * SECOND;

export class Account {
  /**
   * `host`: { transport, now, mono, timers, logger, memory(provider),
   * saveMemory(), interval(), rule(), statusChanged(), presenceChanged(),
   * signedInChanged(), saveSession(provider, value), gladysConnected() }.
   */
  constructor(provider, host) {
    this.provider = provider;
    this.host = host;
    this.logger = host.logger;
    this.now = host.now;
    this.mono = host.mono;
    this.timers = host.timers;
    // Every request counts in the daily budget; a socket error becomes a ServiceError.
    this.transport = async (...args) => {
      if (this.closed) {
        throw new ServiceError('stale', 'stopped');
      }
      this.countRequest();
      try {
        return await host.transport(...args);
      } catch (err) {
        throw networkError(err);
      }
    };
    this.config = { client: '', started: false };
    /** Moved on by every sign-in and every start: an answer of before is dropped. */
    this.epoch = 0;
    this.generation = 0;
    this.timer = undefined;
    this.renewTimer = undefined;
    this.running = undefined;
    this.renewing = undefined;
    this.lastTestAt = -Infinity;
    /** Monotonic time of the start of the last read of this run. */
    this.lastReadMono = -Infinity;
    this.pausedUntil = 0;
    this.pauseKind = undefined;
    /** The tokens (in memory; saved in the configuration of Gladys, never in /data). */
    this.stored = null;
    this.sessionLoaded = false;
    /** Tokens not saved in Gladys yet (a failed setConfig): saved again soon (short timer). */
    this.sessionPending = false;
    /** The writes of the session to Gladys, one at a time (never an older value last). */
    this.sessionChain = Promise.resolve();
    this.sessionTimer = undefined;
    this.sessionFailures = 0;
    /** The container stopped (after shutdown): no timer, no request any more. */
    this.closed = false;
    /** The container is stopping (shutdown started): no renewal leaves any more. */
    this.closing = false;
    this.reset();
  }

  get memory() {
    return this.host.memory(this.provider);
  }

  /** Forgets the state of this run. A 429, 5xx or budget pause stays. */
  reset() {
    /** The last presence read: { presence, at (monotonic) }, in memory only. */
    this.reading = null;
    /** undefined | 'no_client' | 'signed_out' | 'expired' | 'unauthorized' */
    this.blocked = undefined;
    this.offline = false;
    this.connectedOnce = false;
    this.failures = 0;
    this.renewFailures = 0;
    this.renewRetryAt = 0;
    this.renewError = undefined;
    this.cycleStart = -Infinity;
    /** { connected, message } of this account (null: not configured). */
    this.status = null;
    if (!HARD.has(this.pauseKind) || this.now() >= this.pausedUntil) {
      this.pausedUntil = 0;
      this.pauseKind = undefined;
    }
  }

  // --- Lifecycle -------------------------------------------------------------

  /** (Re)starts with the account configuration `config` (src/config.js), /data already read. */
  async start(config) {
    this.stop();
    this.epoch += 1;
    await this.running?.catch(() => {});
    // A renewal in flight ends first: its new renewal token is kept before going on.
    await this.renewing?.catch(() => {});
    const previous = this.config.client;
    this.config = config;
    if (!this.sessionLoaded) {
      // The tokens saved in Gladys are read once; afterwards this process holds the truth.
      this.stored = config.stored ?? null;
      this.sessionLoaded = true;
    }
    if (this.sessionPending) {
      // Not saved before this start (stop() cleared its timer): tried again soon.
      this.scheduleSessionSave();
    }
    if (previous !== config.client) {
      this.lastReadMono = -Infinity;
    }
    this.reset();
    const pause = this.memory.pause;
    if (pause && pause.until > Math.max(this.now(), this.pausedUntil)) {
      this.pausedUntil = pause.until;
      this.pauseKind = pause.kind;
    }
    const { client } = config;
    if (!client) {
      this.blocked = 'no_client';
      // Nothing typed: this service is simply not used (no status line).
      this.setStatus(config.started ? { connected: false, message: T.incomplete } : null);
      return;
    }
    if (this.memory.block?.client === client) {
      // A stop on this very client survives a restart (never retried by itself).
      this.blocked = this.memory.block.kind;
      this.setStatus({ connected: false, message: STOP_MESSAGES[this.blocked] });
      return;
    }
    if (!this.session) {
      this.blocked = 'signed_out';
      this.setStatus({ connected: false, message: T.signedOut });
      return;
    }
    this.scheduleRenew();
    const due = this.dueAt(this.now());
    if (due > this.now()) {
      const success = this.memory.lastSuccessAt;
      // The last read succeeded and nothing failed nor paused since: still connected.
      const connected =
        this.pausedUntil <= this.now() &&
        success !== null &&
        success >= (this.memory.lastCycleAt ?? -Infinity);
      this.connectedOnce = connected;
      this.setStatus({
        connected,
        message: join(
          connected ? T.connected : this.pauseMessage(),
          T.nextRead(age(due - this.now())),
        ),
      });
      this.scheduleAt(due);
      return;
    }
    this.setStatus({ connected: false, message: this.pauseMessage() ?? T.connecting });
    await this.tick();
  }

  stop() {
    this.generation += 1;
    // An exchange or a renewal in flight never schedules anything after it.
    this.epoch += 1;
    this.clearTimer();
    this.clearRenewTimer();
    this.clearSessionTimer();
  }

  /**
   * The container stops (MeetingApp.shutdown): a renewal in flight, then the
   * session writes asked, are waited for (the new renewal token is kept).
   */
  async settle() {
    await this.renewing?.catch(() => {});
    await this.sessionChain;
  }

  clearTimer() {
    if (this.timer) {
      this.timers.clearTimeout(this.timer);
      this.timer = undefined;
    }
  }

  clearRenewTimer() {
    if (this.renewTimer) {
      this.timers.clearTimeout(this.renewTimer);
      this.renewTimer = undefined;
    }
  }

  /** The session of the configured client, or null. */
  get session() {
    const s = this.stored;
    return s && this.config.client && s.client === this.config.client ? s : null;
  }

  /**
   * Counts in "in a meeting, Teams or Zoom": an account with a session, or
   * stopped (renewal or access refused) on its current keys. Its value is
   * then unknown, never left out: only keys emptied (or never authorized)
   * take a service out.
   */
  get signedIn() {
    const client = this.config.client;
    return (
      Boolean(client) &&
      (Boolean(this.session) || this.host.memory(this.provider)?.block?.client === client)
    );
  }

  /** New tokens (or none): in memory at once, then saved in Gladys. */
  async setSession(value) {
    this.stored = value;
    this.sessionPending = true;
    return this.saveSession();
  }

  /**
   * Saves the tokens in the configuration of Gladys; false when it failed
   * (tried again by a short timer). One write at a time: each sends the
   * session of the moment it leaves, so an older value is never written last.
   * `witness`: the current session is written even when Gladys has it (a
   * renewal is sent only once Gladys just accepted a write).
   */
  saveSession({ witness = false } = {}) {
    // A write asked before a stop() still leaves, but its failure arms no timer.
    const generation = this.generation;
    const link = this.sessionChain.then(() => this.writeSession(witness, generation));
    this.sessionChain = link.catch(() => {});
    return link;
  }

  async writeSession(witness, generation) {
    if (this.closed) {
      return !this.sessionPending;
    }
    if (!this.sessionPending && !witness) {
      return true;
    }
    const value = this.stored;
    try {
      await this.host.saveSession(this.provider, value);
    } catch (err) {
      this.logger.warn(`${this.provider}: tokens not saved in Gladys (${err?.status ?? 'error'})`);
      if (this.sessionPending && generation === this.generation) {
        this.scheduleSessionSave();
      }
      return false;
    }
    // A newer session set meanwhile stays to save (its own write follows).
    if (this.stored === value) {
      this.sessionPending = false;
      this.sessionFailures = 0;
      this.clearSessionTimer();
    }
    return true;
  }

  /** A failed save of the session: tried again after 10 s, then at most every minute. */
  scheduleSessionSave() {
    this.clearSessionTimer();
    if (this.closed) {
      return;
    }
    const delay =
      TIMINGS.sessionRetry[Math.min(this.sessionFailures, TIMINGS.sessionRetry.length - 1)];
    this.sessionFailures += 1;
    this.sessionTimer = this.timers.setTimeout(() => {
      this.sessionTimer = undefined;
      this.saveSession().catch(() => {});
    }, delay);
  }

  clearSessionTimer() {
    if (this.sessionTimer) {
      this.timers.clearTimeout(this.sessionTimer);
      this.sessionTimer = undefined;
    }
  }

  get interval() {
    return this.host.interval();
  }

  /**
   * The time of the next read, from one reading of the clock `now`: not
   * sooner than the interval after the last read (kept in /data), nor before
   * the end of a pause or of a renewal wait.
   */
  dueAt(now) {
    return Math.max(
      (this.memory.lastCycleAt ?? -Infinity) + this.interval,
      this.pausedUntil,
      this.renewRetryAt,
      now,
    );
  }

  setStatus(status) {
    this.status = status;
    this.host.statusChanged();
  }

  tokenOptions(timeout) {
    return {
      provider: this.provider,
      clientId: this.config.clientId,
      clientSecret: this.config.clientSecret,
      tenant: this.config.tenant,
      now: this.now(),
      timeout,
    };
  }

  /** One more request today; refused once the daily budget is spent. */
  countRequest() {
    const day = dayOf(this.now());
    const memory = this.memory;
    if (memory.requests.day !== day) {
      memory.requests = { day, count: 0 };
    }
    if (memory.requests.count >= LIMITS.dailyRequests) {
      throw new ServiceError('budget', 'daily number of requests reached');
    }
    memory.requests.count += 1;
  }

  // --- Pauses and stops ----------------------------------------------------------

  /** A 429, a 5xx, a 403 or the daily budget: a pause kept in /data. */
  async pause(err) {
    let wait;
    const now = this.now();
    if (err.kind === 'rate_limited') {
      wait = Math.min(
        TIMINGS.rateLimitMax,
        Math.max(TIMINGS.rateLimitMin, err.retryAfter ?? TIMINGS.rateLimitDefault),
      );
      this.pauseKind = 'rate';
    } else if (err.kind === 'budget') {
      // Until the next UTC day, plus a minute.
      wait = (dayOf(now) + 1) * DAY - now + MINUTE;
      this.pauseKind = 'budget';
    } else if (err.kind === 'forbidden') {
      wait = Math.max(this.interval, TIMINGS.forbidden);
      this.pauseKind = 'forbidden';
    } else {
      const step = TIMINGS.unavailable[Math.min(this.failures, TIMINGS.unavailable.length - 1)];
      wait = Math.max(this.interval, step);
      this.pauseKind = 'unavailable';
      this.offline = true;
    }
    this.pausedUntil = Math.max(this.pausedUntil, now + wait);
    this.memory.pause = { kind: this.pauseKind, until: this.pausedUntil };
    this.memory.lastSuccessAt = null;
    await this.host.saveMemory();
  }

  /** The message of a pause still running. */
  pauseMessage(now = this.now()) {
    const left = this.pausedUntil - now;
    if (left <= 0) {
      return undefined;
    }
    switch (this.pauseKind) {
      case 'rate':
        return T.rateLimited(age(left));
      case 'unavailable':
        return T.unavailable(age(left));
      case 'budget':
        return T.budget(age(left));
      case 'forbidden':
        return T.forbidden(age(left));
      default:
        return undefined;
    }
  }

  /** The stop message of this account (the hint of Connect is added to it, never twice). */
  stopMessage() {
    return STOP_MESSAGES[this.blocked] ?? T.signedOut;
  }

  /**
   * A pause never longer than its kind allows from `now`: a wall clock set
   * back never stretches it (capped at every use, as when read from /data).
   */
  capPause(now) {
    if (this.pauseKind && this.pausedUntil > now) {
      this.pausedUntil = Math.min(this.pausedUntil, now + PAUSES[this.pauseKind]);
    }
    if (this.renewRetryAt > now) {
      this.renewRetryAt = Math.min(this.renewRetryAt, now + TIMINGS.renewRetry.at(-1));
    }
  }

  get hardPaused() {
    return HARD.has(this.pauseKind) && this.now() < this.pausedUntil;
  }

  /** A stop only the user lifts: kept in /data with its client. */
  async block(kind) {
    const wasSignedIn = this.signedIn;
    this.blocked = kind;
    this.memory.lastSuccessAt = null;
    this.reading = null;
    this.clearTimer();
    if (kind === 'expired') {
      // The renewal token is useless now: a new authorization replaces it.
      this.clearRenewTimer();
      await this.setSession(null);
    }
    if (PERSISTED.has(kind)) {
      this.memory.block = { kind, client: this.config.client };
    }
    await this.host.saveMemory();
    this.logger.warn(`${this.provider}: stopped (${kind})`);
    this.setStatus({ connected: false, message: STOP_MESSAGES[kind] });
    if (wasSignedIn !== this.signedIn) {
      this.host.signedInChanged();
    }
    this.host.presenceChanged();
  }

  // --- Token renewal ----------------------------------------------------------

  /**
   * The renewal timer: at the end of the token minus a margin (10 minutes, at
   * most half of its lifetime), never within a minute of the last renewal,
   * nor during a hard pause, nor before the next try after a failure.
   */
  scheduleRenew() {
    this.clearRenewTimer();
    if (this.closed) {
      return;
    }
    const s = this.session;
    if (!s || PERSISTED.has(this.blocked)) {
      return;
    }
    const lifetime = Math.max(0, s.expiresAt - s.obtainedAt);
    const margin = Math.min(TIMINGS.tokenMargin, lifetime / 2);
    const now = this.now();
    let at = Math.max(
      s.expiresAt - margin,
      (this.memory.lastRenewAt ?? -Infinity) + TIMINGS.renewFloor,
      this.renewRetryAt,
    );
    if (HARD.has(this.pauseKind) && now < this.pausedUntil) {
      at = Math.max(at, this.pausedUntil);
    }
    const generation = this.generation;
    this.renewTimer = this.timers.setTimeout(
      () => {
        this.renewTimer = undefined;
        if (generation !== this.generation) {
          return;
        }
        this.renew().catch((err) => this.renewFailed(err, generation));
      },
      Math.max(0, at - now),
    );
  }

  /** A renewal of the timer failed: a pause, a later try, or a stop (refused). */
  async renewFailed(err, generation) {
    if (generation !== this.generation || this.blocked) {
      return;
    }
    this.logger.info(`${this.provider}: token renewal failed (${errorReason(err)})`);
    if (PAUSING.has(err.kind)) {
      await this.pause(err);
      if (generation !== this.generation) {
        return;
      }
      this.setStatus({
        connected: this.connectedOnce && err.kind !== 'unavailable',
        message: this.pauseMessage(),
      });
    } else if (err.kind === 'renew_wait' || err.kind === 'paused') {
      // A renewal ran less than a minute ago, or a pause runs: the timer comes back after it.
    } else if (this.session && this.now() >= this.session.expiresAt) {
      this.setStatus({
        connected: false,
        message: T.renewWaiting(errorReason(err), age(this.renewRetryAt - this.now())),
      });
    }
    this.scheduleRenew();
  }

  /** One renewal at a time: a second call joins the one in flight. */
  renew(timeout) {
    if (!this.renewing) {
      this.renewing = this.doRenew(timeout).finally(() => {
        this.renewing = undefined;
      });
    }
    return this.renewing;
  }

  async doRenew(timeout = TIMINGS.requestTimeout) {
    const s = this.session;
    if (!s || this.blocked) {
      throw new ServiceError('signed_out', 'not signed in');
    }
    const now = this.now();
    if (now < (this.memory.lastRenewAt ?? -Infinity) + TIMINGS.renewFloor) {
      throw new ServiceError('renew_wait', 'renewed less than a minute ago');
    }
    if (this.hardPaused) {
      // Nothing sent: the pause already running is neither extended nor restarted.
      throw new ServiceError('paused', 'paused');
    }
    // Renewal tokens serve once: never a renewal Gladys could not keep. The
    // session is written first (the pending one, or the current one again);
    // refused or Gladys not connected, nothing is sent and the token stays.
    if (!this.host.gladysConnected() || !(await this.saveSession({ witness: true }))) {
      this.renewRetryAt = this.now() + TIMINGS.gladysWait;
      this.renewError = new ServiceError('gladys', 'Gladys cannot keep the tokens');
      throw new ServiceError('renew_wait', 'Gladys cannot keep the tokens');
    }
    if (s !== this.session) {
      throw new ServiceError('renew_wait', 'a new sign-in replaced the session');
    }
    if (this.closing) {
      // The container stops: its token could not be kept (a read in flight got a 401).
      throw new ServiceError('renew_wait', 'the integration is stopping');
    }
    const epoch = this.epoch;
    // Kept before the request: a restart during it never renews within a minute.
    this.memory.lastRenewAt = now;
    await this.host.saveMemory();
    // Only a new session makes the answer useless; a restart on the same
    // session keeps it: the renewal token sent may be consumed, the new one is kept.
    const replaced = () => {
      const current = this.stored;
      return !current || current.client !== s.client || current.refreshToken !== s.refreshToken;
    };
    let token;
    try {
      token = await tokenRequest(
        this.transport,
        { grant_type: 'refresh_token', refresh_token: s.refreshToken },
        this.tokenOptions(timeout),
      );
    } catch (err) {
      if (replaced()) {
        throw new ServiceError('renew_wait', 'a new sign-in replaced the session');
      }
      if (err.kind === 'refused') {
        // Never tried again by itself: only a new authorization clears it.
        await this.block('expired');
      } else if (!PAUSING.has(err.kind)) {
        this.renewRetryAt =
          this.now() +
          TIMINGS.renewRetry[Math.min(this.renewFailures, TIMINGS.renewRetry.length - 1)];
        this.renewFailures += 1;
        this.renewError = err;
      }
      throw err;
    }
    if (replaced()) {
      return;
    }
    // Renewal tokens rotate (Zoom, Microsoft): the new one is kept at once.
    this.renewFailures = 0;
    this.renewRetryAt = 0;
    this.renewError = undefined;
    // Saved in Gladys before it is used: a restart keeps the renewal token
    // the service now expects (the old one may be refused already).
    await this.setSession({
      client: s.client,
      ...token,
      refreshToken: token.refreshToken ?? s.refreshToken,
    });
    if (epoch === this.epoch && !this.closed) {
      this.scheduleRenew();
    }
  }

  /** A usable access token: renewed first when it ended (and a renewal is allowed). */
  async accessToken(timeout) {
    const s = this.session;
    if (!s) {
      throw new ServiceError('signed_out', 'not signed in');
    }
    if (s.accessToken && this.now() < s.expiresAt - TOKEN_SLACK) {
      return s.accessToken;
    }
    if (this.now() < this.renewRetryAt) {
      throw new ServiceError(
        'renew_wait',
        this.renewError ? errorReason(this.renewError) : 'renewal failed',
      );
    }
    await this.renew(timeout);
    const fresh = this.session;
    if (!fresh?.accessToken) {
      throw new ServiceError('signed_out', 'not signed in');
    }
    return fresh.accessToken;
  }

  /**
   * The presence read. A refused token renews once (when a renewal is
   * allowed) and tries again; refused again with a token of this read: a stop.
   */
  async call(timeout) {
    const fn = (token) =>
      readPresence(this.transport, this.provider, { token, now: this.now(), timeout });
    let token = await this.accessToken(timeout);
    try {
      return await fn(token);
    } catch (err) {
      if (err.kind !== 'unauthorized') {
        throw err;
      }
      const current = this.session?.accessToken;
      if (current && current !== token) {
        // Renewed by the timer meanwhile: tried once more with the new token.
        token = current;
      } else {
        if ((this.session?.obtainedAt ?? 0) >= this.cycleStart) {
          throw new ServiceError('unauthorized_block', 'a new access token was refused');
        }
        if (this.now() < this.renewRetryAt) {
          throw new ServiceError(
            'renew_wait',
            this.renewError ? errorReason(this.renewError) : 'renewal failed',
          );
        }
        // Expired early: renewed now (the floor of a minute still holds).
        await this.renew(timeout);
        token = this.session?.accessToken;
        if (!token) {
          throw err;
        }
      }
    }
    try {
      return await fn(token);
    } catch (err) {
      if (err.kind === 'unauthorized') {
        throw new ServiceError('unauthorized_block', 'a new access token was refused');
      }
      throw err;
    }
  }

  // --- Polling ---------------------------------------------------------------

  scheduleNext() {
    this.clearTimer();
    if (this.blocked) {
      return;
    }
    const now = this.now();
    this.scheduleAt(now + Math.max(this.interval, this.pausedUntil - now, this.renewRetryAt - now));
  }

  scheduleAt(at) {
    this.clearTimer();
    if (this.closed) {
      return;
    }
    const delay = at - this.now();
    const generation = this.generation;
    this.timer = this.timers.setTimeout(
      () => {
        this.timer = undefined;
        if (generation === this.generation) {
          this.tick().catch((err) =>
            this.logger.error(`${this.provider}: read failed: ${err.message}`),
          );
        }
      },
      Math.max(0, delay),
    );
  }

  /** One read at a time: a tick during a read joins it. `quick`: for a button. */
  tick({ quick = false } = {}) {
    if (!this.running) {
      this.running = this.run({ quick }).finally(() => {
        this.running = undefined;
      });
    }
    return this.running;
  }

  async run({ quick = false } = {}) {
    const generation = this.generation;
    if (this.blocked || !this.host.ready() || this.closed || this.closing) {
      return;
    }
    // Tokens not saved in Gladys yet: tried again before the read.
    await this.saveSession();
    if (generation !== this.generation) {
      return;
    }
    const now = this.now();
    this.capPause(now);
    if (now < this.pausedUntil) {
      this.scheduleNext();
      return;
    }
    // One read per interval in this run, whatever the wall clock does.
    const mono = this.mono();
    const sinceLast = mono - this.lastReadMono;
    if (sinceLast < this.interval) {
      this.scheduleAt(now + (this.interval - sinceLast));
      return;
    }
    const timeout = quick ? TIMINGS.quickTimeout : TIMINGS.requestTimeout;
    this.cycleStart = now;
    this.lastReadMono = mono;
    // Kept in /data before the request: a restart never reads sooner than the interval.
    this.memory.lastCycleAt = now;
    await this.host.saveMemory();
    if (generation !== this.generation) {
      return;
    }
    try {
      const presence = await this.call(timeout);
      if (generation !== this.generation) {
        return;
      }
      // A stop raised meanwhile (a refused renewal by the timer) is never covered.
      if (!this.blocked) {
        this.reading = { presence, at: this.mono() };
        await this.succeeded();
      }
    } catch (err) {
      if (generation !== this.generation) {
        return;
      }
      await this.failed(err);
    }
    if (generation !== this.generation) {
      return;
    }
    await this.host.saveMemory();
    this.host.presenceChanged();
    this.scheduleNext();
  }

  /** The last reading if fresh: { inMeeting, text }, or null. */
  freshReading(mono = this.mono()) {
    if (!this.reading || this.blocked) {
      return null;
    }
    const limit = Math.max(3 * this.interval, TIMINGS.freshMin);
    if (mono - this.reading.at > limit) {
      return null;
    }
    const { presence } = this.reading;
    return {
      inMeeting: inMeetingOf(this.provider, presence, this.host.rule()),
      text: textOf(this.provider, presence),
    };
  }

  async succeeded() {
    this.failures = 0;
    this.offline = false;
    this.connectedOnce = true;
    this.pauseKind = undefined;
    this.pausedUntil = 0;
    // Kept in /data: a restart before the next read stays connected (cleared by any failure).
    this.memory.lastSuccessAt = this.cycleStart;
    this.memory.pause = null;
    const known = inMeetingOf(this.provider, this.reading.presence, this.host.rule()) !== null;
    this.setStatus({ connected: true, message: known ? T.connected : T.unknownPresence });
  }

  async failed(err) {
    const reason = errorReason(err);
    this.memory.lastSuccessAt = null;
    switch (err.kind) {
      case 'signed_out':
        if (!this.blocked) {
          this.blocked = 'signed_out';
          this.setStatus({ connected: false, message: T.signedOut });
        }
        return;
      case 'refused':
        // The renewal was refused: block('expired') already ran.
        return;
      case 'unauthorized_block':
        await this.block('unauthorized');
        return;
      case 'paused':
        return;
      default:
        break;
    }
    if (this.blocked) {
      return;
    }
    let message;
    const now = this.now();
    if (PAUSING.has(err.kind) || err.kind === 'forbidden') {
      await this.pause(err);
      message = this.pauseMessage();
    } else if (err.kind === 'renew_wait') {
      const wait = Math.max(this.interval, this.renewRetryAt - now);
      this.pauseKind = 'backoff';
      this.pausedUntil = Math.max(this.pausedUntil, now + wait);
      message = T.renewWaiting(reason, age(wait));
    } else {
      // Never read again sooner than the interval.
      const wait = Math.max(
        this.interval,
        TIMINGS.backoff[Math.min(this.failures, TIMINGS.backoff.length - 1)],
      );
      this.pauseKind = 'backoff';
      this.offline = OFFLINE.has(err.kind) || this.offline;
      this.pausedUntil = Math.max(this.pausedUntil, now + wait);
      message = T.unreachable(reason, age(wait));
    }
    if (this.pauseKind === 'backoff') {
      this.memory.pause = {
        kind: 'backoff',
        until: Math.min(this.pausedUntil, now + TIMINGS.backoffMax),
      };
    }
    this.failures += 1;
    const connected =
      !['unavailable', 'forbidden'].includes(err.kind) &&
      this.pauseKind !== 'backoff' &&
      this.connectedOnce;
    this.setStatus({ connected, message });
    this.logger.warn(`${this.provider}: ${reason}`);
  }

  // --- Sign-in -----------------------------------------------------------------

  /** The authorization address of a sign-in begun by the app. */
  authorizeUrl(state, entry) {
    return authorizeUrl({
      provider: this.provider,
      clientId: this.config.clientId,
      tenant: this.config.tenant,
      redirectUri: entry.redirectUri,
      state,
      verifier: entry.verifier,
    });
  }

  /**
   * The code -> a new session, kept in /data; then one read. `{ ok, message }`:
   * answered once the session is in /data; the status and the read follow
   * without being waited for (Gladys acks the OAuth return within 5 s).
   */
  async exchange(code, signIn, { timeout = TIMINGS.quickTimeout, pasted = false, restore }) {
    const fail = (message) => ({ ok: false, message });
    if (!this.config.client || signIn.entry.client !== this.config.client) {
      // The keys changed since Connect: this code belongs to another app.
      return fail(this.config.client ? T.signIn.unknown : T.signIn.noClient);
    }
    if (this.hardPaused) {
      restore();
      return fail(this.pauseMessage());
    }
    const epoch = this.epoch;
    const client = this.config.client;
    let token;
    try {
      token = await tokenRequest(
        this.transport,
        {
          grant_type: 'authorization_code',
          code,
          redirect_uri: signIn.entry.redirectUri,
          code_verifier: signIn.entry.verifier,
        },
        this.tokenOptions(timeout),
      );
    } catch (err) {
      this.logger.warn(`${this.provider}: authorization failed: ${errorReason(err)}`);
      if (err.kind === 'timeout') {
        // The code may have been used meanwhile: a new authorization is the safe way.
        return fail(T.signIn.slow);
      }
      if (err.kind === 'refused') {
        return fail(T.signIn.refused(err.detail ?? 'HTTP 400'));
      }
      if (PAUSING.has(err.kind)) {
        await this.pause(err);
        restore();
        return fail(this.pauseMessage());
      }
      if (!pasted) {
        return fail(T.signIn.retryConnect(errorReason(err)));
      }
      restore();
      return fail(T.signIn.failed(errorReason(err)));
    }
    if (epoch !== this.epoch || client !== this.config.client) {
      return fail(T.signIn.unknown);
    }
    if (!token.refreshToken) {
      return fail(T.signIn.noRefresh);
    }
    const wasSignedIn = this.signedIn;
    // A renewal still running with the previous session must not overwrite this one.
    this.epoch += 1;
    const mine = this.epoch;
    this.memory.block = null;
    this.blocked = undefined;
    this.failures = 0;
    this.renewFailures = 0;
    this.renewRetryAt = 0;
    this.renewError = undefined;
    this.reading = null;
    if (this.pauseKind === 'backoff' || this.pauseKind === 'forbidden') {
      this.pausedUntil = 0;
      this.pauseKind = undefined;
      this.memory.pause = null;
    }
    await this.setSession({ client, ...token });
    await this.host.saveMemory();
    if (mine !== this.epoch) {
      // Stopped or restarted meanwhile: nothing scheduled from here.
      return { ok: true, message: T.signIn.done };
    }
    this.logger.info(`${this.provider}: authorized`);
    if (wasSignedIn !== this.signedIn) {
      this.host.signedInChanged();
    }
    const now = this.now();
    const due = this.dueAt(now);
    const later = due > now ? T.signIn.doneLater(age(due - now)) : null;
    this.afterSignIn(mine, due, later).catch((err) =>
      this.logger.error(`${this.provider}: read failed: ${err.message}`),
    );
    return { ok: true, message: later ?? T.signIn.done };
  }

  async afterSignIn(mine, due, later) {
    if (mine !== this.epoch) {
      return;
    }
    this.scheduleRenew();
    if (later) {
      this.setStatus({ connected: false, message: later });
      this.scheduleAt(due);
      return;
    }
    this.setStatus({ connected: false, message: T.connecting });
    await this.tick({ quick: true });
  }

  // --- Test button -------------------------------------------------------------

  /** The status of this account after a click on Test (a read only when due). */
  async test() {
    if (!this.config.client) {
      return this.config.started ? T.incomplete : null;
    }
    if (this.blocked === 'signed_out' || this.blocked === 'expired') {
      return STOP_MESSAGES[this.blocked];
    }
    this.capPause(this.now());
    const paused = this.pauseMessage();
    if (paused) {
      return paused;
    }
    const mono = this.mono();
    if (mono - this.lastTestAt < TIMINGS.testCooldown) {
      return join(T.test.cooldown, this.status?.message);
    }
    if (this.running) {
      return T.busy;
    }
    const now = this.now();
    const due = Math.max(this.dueAt(now), now + this.interval - (mono - this.lastReadMono));
    if (this.blocked === 'unauthorized') {
      // One read per click: the button is how a refused access is tried again,
      // still never sooner than the interval after the last read.
      if (due > now) {
        return join(T.unauthorized, T.nextRead(age(due - now)));
      }
      this.lastTestAt = mono;
      this.blocked = undefined;
      this.memory.block = null;
      await this.host.saveMemory();
      this.scheduleRenew();
      await this.tick({ quick: true });
      return this.status?.message ?? T.signedOut;
    }
    if (due > now) {
      return join(this.status?.message, T.nextRead(age(due - now)));
    }
    this.lastTestAt = mono;
    await this.tick({ quick: true });
    return this.status?.message ?? T.signedOut;
  }
}
