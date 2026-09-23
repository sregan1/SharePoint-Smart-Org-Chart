// Shared HTTP helpers for Graph and SharePoint REST calls: status/Retry-After
// extraction and a retry wrapper with exponential backoff for throttling
// (429) and transient server errors.

export interface IHttpError extends Error {
  status?: number;
  retryAfterMs?: number;
}

const MAX_RETRY_DELAY_MS = 30_000;

export const sleep = (ms: number): Promise<void> => new Promise(r => setTimeout(r, ms));

export function createHttpError(message: string, status: number, retryAfterHeader?: string | null): IHttpError {
  const err = new Error(message) as IHttpError;
  err.status = status;
  const retryAfter = parseRetryAfter(retryAfterHeader);
  if (retryAfter !== undefined) err.retryAfterMs = retryAfter;
  return err;
}

// Retry-After is either a number of seconds or an HTTP date
export function parseRetryAfter(value: string | null | undefined): number | undefined {
  if (!value) return undefined;
  const secs = Number(value);
  if (!isNaN(secs)) return Math.max(0, secs * 1000);
  const date = Date.parse(value);
  if (!isNaN(date)) return Math.max(0, date - Date.now());
  return undefined;
}

// Works for our own IHttpError (status) and the Graph SDK's GraphError (statusCode)
export function getErrorStatus(err: unknown): number | undefined {
  if (!err || typeof err !== 'object') return undefined;
  const e = err as { status?: unknown; statusCode?: unknown };
  if (typeof e.status === 'number') return e.status;
  if (typeof e.statusCode === 'number') return e.statusCode;
  return undefined;
}

export function getRetryAfterMs(err: unknown): number | undefined {
  if (!err || typeof err !== 'object') return undefined;
  const e = err as { retryAfterMs?: unknown; headers?: unknown };
  if (typeof e.retryAfterMs === 'number') return e.retryAfterMs;
  const h = e.headers as { get?: (name: string) => string | null } | { [k: string]: string } | undefined;
  if (!h) return undefined;
  if (typeof (h as { get?: unknown }).get === 'function') {
    return parseRetryAfter((h as { get: (name: string) => string | null }).get('Retry-After'));
  }
  const bag = h as { [k: string]: string };
  return parseRetryAfter(bag['Retry-After'] || bag['retry-after']);
}

// Throttling, transient server errors, and network failures (no status, or
// the Graph SDK's -1) are worth retrying; 4xx client errors are not.
export function isRetryableStatus(status: number | undefined): boolean {
  if (status === undefined || status <= 0) return true;
  return status === 408 || status === 429 || status === 500 ||
         status === 502 || status === 503 || status === 504;
}

// Runs fn, retrying retryable failures with exponential backoff (with jitter),
// honoring Retry-After when the server supplies one.
export async function withRetry<T>(fn: () => Promise<T>, maxAttempts = 4, baseDelayMs = 1000): Promise<T> {
  let attempt = 0;
  let lastErr: unknown;
  while (attempt < maxAttempts) {
    attempt++;
    try {
      return await fn();
    } catch (err) {
      lastErr = err;
      if (attempt >= maxAttempts || !isRetryableStatus(getErrorStatus(err))) throw err;
      const retryAfter = getRetryAfterMs(err);
      const backoff = baseDelayMs * Math.pow(2, attempt - 1) * (0.75 + Math.random() * 0.5);
      await sleep(Math.min(MAX_RETRY_DELAY_MS, retryAfter !== undefined ? Math.max(retryAfter, 250) : backoff));
    }
  }
  throw lastErr;
}
