import { createHash } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import type { OutgoingHttpHeaders, RequestOptions } from 'node:http';

type CancellationTokenLike = {
  createPromise: <T>(
    callback: (
      resolve: (value: T | PromiseLike<T>) => void,
      reject: (error: Error) => void,
      onCancel: (callback: () => void) => void,
    ) => void,
  ) => Promise<T>;
};

type FetchDownloadOptions = {
  headers?: OutgoingHttpHeaders | null;
  sha2?: string | null;
  sha512?: string | null;
  cancellationToken: CancellationTokenLike;
  timeout?: number;
};

export type FetchHttpExecutor = {
  request: (
    options: RequestOptions,
    cancellationToken?: CancellationTokenLike,
    data?: Record<string, unknown> | null,
  ) => Promise<string | null>;
  download: (url: URL, destination: string, options: FetchDownloadOptions) => Promise<string>;
  downloadToBuffer: (url: URL, options: FetchDownloadOptions) => Promise<Buffer>;
};

type FetchImpl = (url: string, init?: RequestInit) => Promise<Response>;
const DEFAULT_DOWNLOAD_TIMEOUT_MS = 120_000;

function requestOptionsToUrl(options: RequestOptions): string {
  const protocol = options.protocol ?? 'https:';
  const hostname = options.hostname ?? options.host;
  if (!hostname) throw new Error('Updater request is missing a hostname.');
  const port = options.port ? `:${options.port}` : '';
  const requestPath = options.path ?? '/';
  return `${protocol}//${hostname}${port}${requestPath}`;
}

function toHeaders(headers: RequestOptions['headers'] | OutgoingHttpHeaders | null | undefined) {
  const result = new Headers();
  if (Array.isArray(headers)) {
    for (let index = 0; index < headers.length; index += 2) {
      const name = headers[index];
      const value = headers[index + 1];
      if (name !== undefined && value !== undefined) {
        result.append(String(name), String(value));
      }
    }
    return result;
  }
  for (const [name, value] of Object.entries(headers ?? {})) {
    if (value === undefined || value === null) continue;
    const values = Array.isArray(value) ? value : [value];
    for (const item of values) {
      result.append(name, String(item));
    }
  }
  return result;
}

function runWithCancellation<T>(
  operation: (signal: AbortSignal) => Promise<T>,
  cancellationToken?: CancellationTokenLike,
  timeoutMs?: number,
): Promise<T> {
  const run = (
    resolve: (value: T | PromiseLike<T>) => void,
    reject: (error: Error) => void,
    onCancel: (callback: () => void) => void,
  ) => {
    const controller = new AbortController();
    const timeout =
      typeof timeoutMs === 'number' && timeoutMs > 0
        ? setTimeout(() => controller.abort(), timeoutMs)
        : null;
    onCancel(() => {
      controller.abort();
    });
    operation(controller.signal)
      .then(resolve, reject)
      .finally(() => {
        if (timeout) clearTimeout(timeout);
      });
  };

  if (cancellationToken) {
    return cancellationToken.createPromise<T>(run);
  }
  return new Promise<T>((resolve, reject) => run(resolve, reject, () => {}));
}

async function fetchBuffer(
  fetchImpl: FetchImpl,
  url: string,
  init: RequestInit,
  cancellationToken?: CancellationTokenLike,
  timeoutMs?: number,
): Promise<Buffer> {
  const response = await runWithCancellation(
    (signal) => fetchImpl(url, { ...init, signal }),
    cancellationToken,
    timeoutMs,
  );
  if (!response.ok) {
    throw new Error(`Updater request failed with ${response.status}`);
  }
  return Buffer.from(await response.arrayBuffer());
}

function verifyDownloadedData(data: Buffer, downloadOptions: FetchDownloadOptions) {
  if (downloadOptions.sha512) {
    const actual = createHash('sha512').update(data).digest('base64');
    if (actual !== downloadOptions.sha512) {
      throw new Error(`sha512 mismatch: expected ${downloadOptions.sha512}, got ${actual}`);
    }
  }
  if (downloadOptions.sha2) {
    const actual = createHash('sha256').update(data).digest('hex');
    if (actual !== downloadOptions.sha2.toLowerCase()) {
      throw new Error(`sha2 mismatch: expected ${downloadOptions.sha2}, got ${actual}`);
    }
  }
}

export function createFetchHttpExecutor(
  options: {
    fetch?: FetchImpl;
    mkdir?: (targetPath: string) => Promise<unknown>;
    writeFile?: (targetPath: string, data: Buffer) => Promise<unknown>;
    downloadTimeoutMs?: number;
  } = {},
): FetchHttpExecutor {
  const fetchImpl = options.fetch ?? globalThis.fetch.bind(globalThis);
  const mkdir =
    options.mkdir ?? ((targetPath: string) => fs.promises.mkdir(targetPath, { recursive: true }));
  const writeFile =
    options.writeFile ??
    ((targetPath: string, data: Buffer) => fs.promises.writeFile(targetPath, data));
  const downloadTimeoutMs = options.downloadTimeoutMs ?? DEFAULT_DOWNLOAD_TIMEOUT_MS;

  return {
    async request(requestOptions, cancellationToken, data): Promise<string | null> {
      const headers = toHeaders(requestOptions.headers);
      const body = data ? JSON.stringify(data) : undefined;
      const result = await fetchBuffer(
        fetchImpl,
        requestOptionsToUrl(requestOptions),
        {
          method: requestOptions.method ?? (body ? 'POST' : 'GET'),
          headers,
          body,
          redirect: 'follow',
        },
        cancellationToken,
        requestOptions.timeout,
      );
      return result.length === 0 ? null : result.toString('utf8');
    },
    async download(url, destination, downloadOptions): Promise<string> {
      await mkdir(path.dirname(destination));
      const data = await fetchBuffer(
        fetchImpl,
        url.href,
        {
          headers: toHeaders(downloadOptions.headers),
          redirect: 'follow',
        },
        downloadOptions.cancellationToken,
        downloadOptions.timeout ?? downloadTimeoutMs,
      );
      verifyDownloadedData(data, downloadOptions);
      await writeFile(destination, data);
      return destination;
    },
    async downloadToBuffer(url, downloadOptions): Promise<Buffer> {
      const data = await fetchBuffer(
        fetchImpl,
        url.href,
        {
          headers: toHeaders(downloadOptions.headers),
          redirect: 'follow',
        },
        downloadOptions.cancellationToken,
        downloadOptions.timeout ?? downloadTimeoutMs,
      );
      verifyDownloadedData(data, downloadOptions);
      return data;
    },
  };
}
