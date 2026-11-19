// -----------------------------------------------------------------------------
// The only network access: HTTPS requests to Microsoft and Zoom, and the errors they end
// with. The body read is bounded, the promise settles once, nothing secret
// ever goes into an error message: a text of the services is
// only read to classify the error, never shown nor logged (only a safe code).
// -----------------------------------------------------------------------------

import https from 'node:https';
import { LIMITS, TIMINGS } from './constants.js';

/**
 * A failed call to Microsoft or Zoom. `kind`:
 * - `refused`: the token endpoint refused a code or the renewal token;
 * - `unauthorized`: the API refused the access token (401);
 * - `forbidden`: the API refused the presence (403: consent, account type);
 * - `rate_limited`: 429, `retryAfter` in ms when given;
 * - `unavailable`: 5xx;
 * - `budget`: the daily number of requests of this account is reached;
 * - `network`, `timeout`: no answer;
 * - `api_error`: an answer without the expected content; `bad_response`,
 *   `size`: not the expected answer; `http`: any other status.
 */
export class ServiceError extends Error {
  constructor(kind, message, { status, retryAfter, detail } = {}) {
    super(message);
    this.name = 'ServiceError';
    this.kind = kind;
    this.status = status;
    this.retryAfter = retryAfter;
    this.detail = detail;
  }
}

const NETWORK_REASONS = {
  ECONNREFUSED: 'connection refused',
  ECONNRESET: 'connection reset',
  EHOSTUNREACH: 'host unreachable',
  ENETUNREACH: 'network unreachable',
  ENOTFOUND: 'unknown host name',
  EAI_AGAIN: 'host name not resolved',
  ETIMEDOUT: 'connection timed out',
  EPIPE: 'connection closed',
};

/** A socket error -> ServiceError (never the request: it carries tokens). */
export function networkError(err) {
  if (err instanceof ServiceError) {
    return err;
  }
  return new ServiceError('network', NETWORK_REASONS[err?.code] ?? 'network error');
}

/** A short reason for a log or a status message: no secret. */
export function errorReason(err) {
  if (err instanceof ServiceError) {
    const code = err.detail ? ` ${err.detail}` : '';
    return err.status ? `${err.message}${code} (HTTP ${err.status})` : `${err.message}${code}`;
  }
  return 'unexpected error';
}

/** A code of a service made safe for a message: letters, digits, `_.-`, 40 at most. */
export const safeCode = (value) => {
  const text = typeof value === 'number' ? String(value) : value;
  return typeof text === 'string' && /^[A-Za-z0-9_.-]{1,40}$/.test(text) ? text : undefined;
};

/**
 * One HTTPS request; resolves `{ status, headers, body }` whatever the status.
 * The chunks are kept in a list and joined once (linear); the total size is
 * bounded. The tests replace it (the app takes a `transport`).
 */
export function httpsTransport(
  url,
  {
    method = 'GET',
    headers = {},
    body,
    timeout = TIMINGS.requestTimeout,
    maxBytes = LIMITS.responseBytes,
    // Tests only: an agent trusting the certificate of a local server.
    agent,
  } = {},
) {
  return new Promise((resolve, reject) => {
    const target = new URL(url);
    if (target.protocol !== 'https:') {
      reject(new ServiceError('network', 'only HTTPS is used'));
      return;
    }
    let timer;
    let settled = false;
    const settle = (fn, value) => {
      if (!settled) {
        settled = true;
        clearTimeout(timer);
        fn(value);
      }
    };
    const req = https.request(target, { method, headers, agent }, (res) => {
      const chunks = [];
      let size = 0;
      res.on('data', (chunk) => {
        size += chunk.length;
        if (size > maxBytes) {
          settle(reject, new ServiceError('size', 'answer too large'));
          req.destroy();
          return;
        }
        chunks.push(chunk);
      });
      res.on('end', () =>
        settle(resolve, {
          status: res.statusCode,
          headers: res.headers,
          body: Buffer.concat(chunks).toString('utf8'),
        }),
      );
      res.on('error', (err) => settle(reject, networkError(err)));
    });
    timer = setTimeout(
      () => {
        settle(
          reject,
          new ServiceError('timeout', `no answer within ${Math.round(timeout / 1000)} s`),
        );
        req.destroy();
      },
      Math.max(1, timeout),
    );
    req.on('error', (err) => settle(reject, networkError(err)));
    req.end(body);
  });
}

/**
 * Retry-After (seconds or an HTTP date) -> milliseconds, undefined when
 * absent. A date is compared with the `Date` header of the same answer
 * (the service's clock), `now` only without a readable one.
 */
export function parseRetryAfter(value, now, serverDate) {
  if (value === undefined || value === null || value === '') {
    return undefined;
  }
  const seconds = Number(value);
  if (Number.isFinite(seconds)) {
    return Math.max(0, seconds * 1000);
  }
  const date = Date.parse(value);
  if (Number.isNaN(date)) {
    return undefined;
  }
  const reference = Date.parse(serverDate ?? '');
  return Math.max(0, date - (Number.isNaN(reference) ? now : reference));
}

/**
 * The JSON of a body, `undefined` when empty; `bad_response` when it is not
 * JSON. Keys that could reach a prototype are dropped.
 */
export function parseJson(body) {
  if (body === undefined || body === null || String(body).trim() === '') {
    return undefined;
  }
  try {
    return JSON.parse(body, (key, value) =>
      key === '__proto__' || key === 'constructor' || key === 'prototype' ? undefined : value,
    );
  } catch {
    throw new ServiceError('bad_response', 'the service sent an answer that is not JSON');
  }
}

/** The transient errors: 429, 5xx. Undefined otherwise. */
export function transientError({ status, headers = {} }, now) {
  if (status === 429) {
    return new ServiceError('rate_limited', 'too many requests', {
      status,
      retryAfter: parseRetryAfter(headers['retry-after'], now, headers.date),
    });
  }
  if (status >= 500) {
    return new ServiceError('unavailable', 'the service is unavailable', { status });
  }
  return undefined;
}

/** The error of an API answer by its status (the body is never shown), or undefined for 200. */
export function apiError(res, now) {
  const transient = transientError(res, now);
  if (transient) {
    return transient;
  }
  if (res.status === 200) {
    return undefined;
  }
  if (res.status === 401) {
    return new ServiceError('unauthorized', 'the service refused the access token', {
      status: 401,
    });
  }
  if (res.status === 403) {
    return new ServiceError('forbidden', 'the service refused the presence', { status: 403 });
  }
  return new ServiceError('http', 'unexpected answer of the service', { status: res.status });
}
