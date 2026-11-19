// -----------------------------------------------------------------------------
// The integration: two accounts (src/account.js), Microsoft Teams presence
// by Microsoft Graph and Zoom presence, one Gladys device and its states
// (read only), the transport badge, the connection status and the
// configuration buttons. index.js only wires the SDK handlers to it.
//
// /data is read once at start; unreadable (not missing), it is read again
// every 10 s, one read at a time, and nothing is sent nor written meanwhile
// (never a session or a stop lost by writing over it).
//
// The states: published only when they change (publishChangedStates), from
// fresh readings only (never older than 3 intervals): an unknown value is
// never published, never as 0. A reconnection to Gladys resends the status,
// the Discovery list and the badge, never a read.
// -----------------------------------------------------------------------------

import { DEVICE_TRANSPORTS, createLogger } from '@gladysassistant/integration-sdk';
import { performance } from 'node:perf_hooks';
import { Account } from './account.js';
import { normalizeConfig, sessionKey } from './config.js';
import { LIMITS, OAUTH_FIELDS, PROVIDERS, SECOND, TIMINGS } from './constants.js';
import { buildDevice, deviceIds, deviceStates } from './devices.js';
import { httpsTransport, safeCode } from './http.js';
import { T, age, join, prefixed } from './i18n.js';
import { SignIns, baseOf, parseRedirect, unwrapState } from './oauth.js';
import { fileStore, normalizeMemory } from './store.js';

/** Unreadable /data: read again after this. */
const STORE_RETRY = 10 * SECOND;
/** getConfig() failed: read again after 5 s, doubled up to a minute. */
const CONFIG_RETRY = 5 * SECOND;
const CONFIG_RETRY_MAX = 60 * SECOND;

/** One message in both languages, for an error thrown to Gladys. */
const thrown = (message) => new Error(`${message.en} / ${message.fr}`);

export class MeetingApp {
  constructor({
    gladys,
    logger = createLogger({ name: 'meeting-status' }),
    transport = httpsTransport,
    now = Date.now,
    mono = () => performance.now(),
    timers = globalThis,
    store = fileStore(),
  }) {
    this.gladys = gladys;
    this.logger = logger;
    this.now = now;
    this.mono = mono;
    this.timers = timers;
    this.store = store;
    this.signIns = new SignIns({ now: mono });
    this.config = normalizeConfig();
    this.memory = null;
    this.loaded = false;
    this.storeTimer = undefined;
    this.loading = undefined;
    this.started = false;
    this.lastStatus = undefined;
    this.published = undefined;
    this.badge = undefined;
    this.statusQueued = false;
    this.storeBroken = false;
    this.rawConfig = {};
    this.configTimer = undefined;
    this.configTry = 0;
    this.connection = 0;
    const host = {
      transport,
      now,
      mono,
      timers,
      logger,
      memory: (provider) => this.memory[provider],
      saveMemory: () => this.saveMemory(),
      interval: () => this.config.pollSeconds * SECOND,
      rule: () => this.config.rule,
      ready: () => this.loaded,
      statusChanged: () => this.statusChanged(),
      presenceChanged: () => this.presenceChanged(),
      signedInChanged: () => this.publishDevices(),
      saveSession: (provider, value) =>
        this.gladys.setConfig({ [sessionKey(provider)]: value ?? '' }),
    };
    this.accounts = Object.fromEntries(PROVIDERS.map((p) => [p, new Account(p, host)]));
  }

  // --- Lifecycle -------------------------------------------------------------

  async start(raw = {}) {
    this.stop();
    this.config = normalizeConfig(raw);
    this.signIns.clear();
    if (!this.loaded) {
      await this.loadStore();
      if (!this.loaded) {
        return;
      }
    }
    await this.migrateTokens();
    await Promise.all(PROVIDERS.map((p) => this.accounts[p].start(this.config.accounts[p])));
    await this.publishDevices();
    await this.updateStatus();
  }

  /** The container stops: every timer cleared, the configuration retries too. */
  shutdown() {
    this.connection += 1;
    for (const account of Object.values(this.accounts)) {
      account.closed = true;
    }
    if (this.configTimer) {
      this.timers.clearTimeout(this.configTimer);
      this.configTimer = undefined;
    }
    this.stop();
  }

  /**
   * Tokens an older /data held: moved to the configuration of Gladys (when
   * Gladys holds none for that service), then removed from /data once Gladys
   * kept them (a failed save keeps them in /data for the next start).
   */
  async migrateTokens() {
    let changed = false;
    for (const p of PROVIDERS) {
      const legacy = this.memory[p].legacySession;
      if (!legacy) {
        continue;
      }
      const account = this.accounts[p];
      let kept = true;
      if (!account.sessionLoaded && !this.config.accounts[p].stored) {
        account.stored = legacy;
        account.sessionLoaded = true;
        account.sessionPending = true;
        kept = await account.saveSession();
      }
      if (kept) {
        delete this.memory[p].legacySession;
        changed = true;
      }
    }
    if (changed) {
      await this.saveMemory();
    }
  }

  stop() {
    this.clearStoreTimer();
    for (const account of Object.values(this.accounts)) {
      account.stop();
    }
  }

  clearStoreTimer() {
    if (this.storeTimer) {
      this.timers.clearTimeout(this.storeTimer);
      this.storeTimer = undefined;
    }
  }

  /** Reads /data; unreadable: the status says it, read again in 10 s (one read at a time). */
  async loadStore() {
    if (this.loading) {
      return this.loading;
    }
    this.loading = (async () => {
      try {
        const raw = await this.store.load();
        this.memory = normalizeMemory(raw, this.now());
        this.loaded = true;
        this.storeBroken = false;
      } catch (err) {
        this.storeBroken = true;
        this.logger.warn(`Reading /data failed: ${err?.code ?? err?.name ?? 'error'}`);
        await this.setStatus(false, T.storeUnreadable);
        this.clearStoreTimer();
        this.storeTimer = this.timers.setTimeout(() => {
          this.storeTimer = undefined;
          this.start(this.rawConfig).catch((e) => this.logger.error(`Start failed: ${e.message}`));
        }, STORE_RETRY);
      }
    })().finally(() => {
      this.loading = undefined;
    });
    return this.loading;
  }

  /**
   * The SDK says connected: the configuration is read (getConfig), and read
   * again on failure every 5 s, 10 s… up to a minute, without end, the
   * status saying so; a newer connection cancels the older retries.
   */
  async onConnected(getConfig) {
    this.connection += 1;
    const connection = this.connection;
    if (this.configTimer) {
      this.timers.clearTimeout(this.configTimer);
      this.configTimer = undefined;
    }
    this.configTry = 0;
    for (;;) {
      let raw;
      try {
        raw = await getConfig();
      } catch (err) {
        if (connection !== this.connection) {
          return;
        }
        const wait = Math.min(CONFIG_RETRY_MAX, CONFIG_RETRY * 2 ** this.configTry);
        this.configTry += 1;
        this.logger.warn(`Reading the configuration failed: ${err?.message ?? 'error'}`);
        await this.setStatus(false, T.configUnreadable(age(wait)));
        await new Promise((resolve) => {
          this.configTimer = this.timers.setTimeout(() => {
            this.configTimer = undefined;
            resolve();
          }, wait);
        });
        if (connection !== this.connection) {
          return;
        }
        continue;
      }
      if (connection !== this.connection) {
        return;
      }
      // A status set while the configuration was unreadable is replaced.
      if (this.configTry > 0) {
        this.lastStatus = undefined;
      }
      await this.connected(raw ?? {});
      if (this.configTry > 0) {
        await this.updateStatus();
      }
      return;
    }
  }

  /**
   * Connected to Gladys. The first time, start; after a reconnection, resend
   * what Gladys keeps in memory only (status, Discovery list, badge). The
   * pauses, the stops and the sessions stay; a read is never started by it.
   */
  async connected(raw = {}) {
    this.rawConfig = raw;
    if (!this.started) {
      this.started = true;
      await this.start(raw);
      return;
    }
    await this.configUpdated(raw);
    this.published = undefined;
    this.badge = undefined;
    const status = this.lastStatus;
    this.lastStatus = undefined;
    if (status) {
      const [connected, message] = JSON.parse(status);
      await this.setStatus(connected, message ?? undefined);
    }
    await this.publishDevices();
    await this.publishBadge();
  }

  /** New keys restart that account; a new interval reschedules; a new rule or language republishes. */
  async configUpdated(raw = {}) {
    this.rawConfig = raw;
    const next = normalizeConfig(raw);
    const previous = this.config;
    this.config = next;
    if (!this.loaded) {
      return this.loadStore().then(() => (this.loaded ? this.start(raw) : undefined));
    }
    for (const p of PROVIDERS) {
      const a = previous.accounts[p];
      const b = next.accounts[p];
      if (a.client !== b.client || a.clientSecret !== b.clientSecret || a.started !== b.started) {
        this.signIns.clear(p);
        await this.accounts[p].start(b);
      } else if (next.pollSeconds !== previous.pollSeconds && this.accounts[p].timer) {
        const account = this.accounts[p];
        account.scheduleAt(account.dueAt(this.now()));
      }
    }
    if (next.language !== previous.language) {
      this.published = undefined;
    }
    await this.publishDevices();
    // A rule, an interval or the services counted may have changed: the values again.
    await this.presenceChanged();
    await this.updateStatus();
    return undefined;
  }

  async saveMemory() {
    if (!this.loaded) {
      // Never written over a file that could not be read.
      return;
    }
    await this.store
      .save(this.memory)
      .catch((err) =>
        this.logger.warn(`Saving the state to /data failed: ${err?.code ?? 'error'}`),
      );
  }

  // --- Status ----------------------------------------------------------------

  async setStatus(connected, message) {
    const key = JSON.stringify([connected, message ?? null]);
    if (key === this.lastStatus) {
      return;
    }
    this.lastStatus = key;
    await this.gladys
      .setConnectionStatus(connected, message)
      .catch((err) => this.logger.warn(`setConnectionStatus failed: ${err.message}`));
  }

  /** An account changed its status: one update, after the current turn. */
  statusChanged() {
    if (this.statusQueued) {
      return;
    }
    this.statusQueued = true;
    queueMicrotask(() => {
      this.statusQueued = false;
      this.updateStatus().catch(() => {});
      this.publishBadge().catch(() => {});
    });
  }

  /** The status of the integration: one line per used service. */
  async updateStatus() {
    if (!this.loaded) {
      return;
    }
    const lines = PROVIDERS.map((p) => [p, this.accounts[p].status]).filter(([, s]) => s);
    if (lines.length === 0) {
      await this.setStatus(false, T.nothing);
      return;
    }
    const connected = lines.every(([, s]) => s.connected);
    await this.setStatus(
      connected,
      join(...lines.map(([p, s]) => (s.message ? prefixed(p, s.message) : null))),
    );
  }

  // --- Gladys ----------------------------------------------------------------

  get signedIn() {
    return PROVIDERS.filter((p) => this.accounts[p].signedIn);
  }

  /** The device once an account was authorized (never for keys typed only). */
  get hasDevice() {
    return PROVIDERS.some(
      (p) =>
        this.accounts[p].session ||
        (this.memory?.[p]?.block && this.memory[p].block.client === this.accounts[p].config.client),
    );
  }

  createdDevice() {
    const id = deviceIds(this.gladys).device;
    return (this.gladys.devices ?? []).find((d) => d.external_id === id);
  }

  async publishDevices() {
    if (!this.loaded || !this.hasDevice) {
      return;
    }
    const list = [buildDevice(this.gladys, this.config.language)];
    const json = JSON.stringify(list);
    if (json === this.published) {
      return;
    }
    try {
      await this.gladys.publishDiscoveredDevices(list);
      this.published = json;
    } catch (err) {
      this.logger.warn(`Publishing the discovered devices failed: ${err.message}`);
    }
  }

  /** The fresh readings, by service. */
  freshReadings() {
    const mono = this.mono();
    const fresh = {};
    for (const p of PROVIDERS) {
      const reading = this.accounts[p].freshReading(mono);
      if (reading) {
        fresh[p] = reading;
      }
    }
    return fresh;
  }

  /** A reading changed (or a stop): the fresh states, the changed ones only. */
  async presenceChanged() {
    await this.publishBadge();
    const device = this.createdDevice();
    if (!device) {
      return;
    }
    const created = new Set((device.features ?? []).map((f) => f.external_id));
    const states = deviceStates(this.gladys, this.freshReadings(), this.signedIn).filter((s) =>
      created.has(s.device_feature_external_id),
    );
    if (states.length === 0) {
      return;
    }
    try {
      await this.gladys.publishChangedStates(states);
    } catch (err) {
      this.logger.warn(`Sending to Gladys failed: ${err.message}`);
    }
  }

  /** The badge of the created device: cloud, or unreachable (no account read lately). */
  async publishBadge() {
    const device = this.createdDevice();
    if (!device) {
      return;
    }
    const ok = PROVIDERS.some((p) => {
      const a = this.accounts[p];
      return a.signedIn && !a.blocked && !a.offline && a.connectedOnce;
    });
    const transport = ok ? DEVICE_TRANSPORTS.CLOUD : DEVICE_TRANSPORTS.UNREACHABLE;
    if (this.badge === transport) {
      return;
    }
    try {
      await this.gladys.publishTransports([{ external_id: device.external_id, transport }]);
      this.badge = transport;
    } catch (err) {
      this.logger.warn(`Transport badge failed: ${err.message}`);
    }
  }

  /** Created or updated (the SDK forgot its states): the fresh values again. */
  async deviceCreated() {
    this.badge = undefined;
    await this.presenceChanged();
  }

  async deviceUpdated() {
    this.badge = undefined;
    await this.presenceChanged();
  }

  async deviceDeleted() {
    this.badge = undefined;
  }

  /** The Discovery tab asks for a scan: the device known now, never a read. */
  async scan() {
    this.published = undefined;
    await this.publishDevices();
  }

  // --- Sign-in -----------------------------------------------------------------

  providerOf(key) {
    return PROVIDERS.find((p) => OAUTH_FIELDS[p] === key);
  }

  /** A Connect button: the authorization address of its service (a new state). */
  async oauthAuthorizeUrl(key, redirectUri) {
    const provider = this.providerOf(key);
    if (!provider) {
      throw thrown(T.signIn.unknownField);
    }
    if (!this.loaded) {
      throw thrown(this.storeBroken ? T.storeUnreadable : T.starting);
    }
    const account = this.accounts[provider];
    if (!account.config.client) {
      throw thrown(T.signIn.noClient);
    }
    if (account.hardPaused) {
      throw thrown(account.pauseMessage());
    }
    const redirect = this.config.redirectUri || redirectUri;
    if (!redirect) {
      throw thrown(T.signIn.noRedirect);
    }
    if (account.blocked && !this.config.redirectUri) {
      // Not connected: the status shows the return address to register in the app.
      account.setStatus({
        connected: false,
        message: join(account.stopMessage(), T.signIn.redirectHint(baseOf(redirect))),
      });
    }
    const { state, entry } = this.signIns.begin({
      provider,
      client: account.config.client,
      redirectUri: redirect,
    });
    return account.authorizeUrl(state, entry);
  }

  /** The service redirected to Gladys: the code of an authorization started here. */
  async oauthCallback(key, { code, state } = {}) {
    const provider = this.providerOf(key);
    if (typeof code !== 'string' || !code || code.length > LIMITS.codeLength) {
      throw thrown(T.signIn.noCode);
    }
    const inner = this.signIns.peek(state) ? state : unwrapState(state);
    // Only for the service of its Connect button (another one leaves it untouched).
    if (!provider || this.signIns.peek(inner)?.provider !== provider) {
      throw thrown(T.signIn.unknown);
    }
    // The state serves once.
    const signIn = this.signIns.take(inner);
    if (!signIn) {
      throw thrown(T.signIn.unknown);
    }
    const result = await this.accounts[provider].exchange(code, signIn, {
      timeout: TIMINGS.callbackTimeout,
      restore: () => this.signIns.restore(signIn),
    });
    if (!result.ok) {
      throw thrown(result.message);
    }
  }

  /** The address pasted by the user (redirect URI other than Gladys'). */
  async finishSignIn(fields) {
    const parsed = parseRedirect(fields?.url);
    if (!parsed) {
      return T.signIn.noCode;
    }
    if (parsed.error) {
      // Only a safe code: a text of the pasted address is never shown back.
      return T.signIn.cancelled(safeCode(parsed.error) ?? 'error');
    }
    if (!parsed.state || !parsed.code) {
      return T.signIn.noCode;
    }
    // The state the provider sent back may be the one wrapped by the Gladys front.
    const state = this.signIns.peek(parsed.state) ? parsed.state : unwrapState(parsed.state);
    const pending = this.signIns.peek(state);
    if (!pending) {
      return T.signIn.unknown;
    }
    if (parsed.base !== null && parsed.base !== baseOf(pending.redirectUri)) {
      return T.signIn.wrongAddress;
    }
    const signIn = this.signIns.take(state);
    if (!signIn) {
      return T.signIn.unknown;
    }
    const result = await this.accounts[signIn.entry.provider].exchange(parsed.code, signIn, {
      timeout: TIMINGS.quickTimeout,
      pasted: true,
      restore: () => this.signIns.restore(signIn),
    });
    return prefixed(signIn.entry.provider, result.message);
  }

  // --- Configuration buttons ----------------------------------------------------

  async testConnection() {
    if (!this.loaded) {
      return this.storeBroken ? T.storeUnreadable : T.starting;
    }
    const lines = [];
    // One after the other: each test waits for its own read (bounded by quickTimeout).
    for (const p of PROVIDERS) {
      const message = await this.accounts[p].test();
      if (message) {
        lines.push(prefixed(p, message));
      }
    }
    return lines.length ? join(...lines) : T.nothing;
  }
}
