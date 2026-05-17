import { execFile as defaultExecFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import type { RequestOptions, OutgoingHttpHeaders } from 'node:http';

export type CurlExecFile = (
  file: string,
  args: readonly string[],
  options: {
    encoding: 'utf8' | 'buffer';
    maxBuffer?: number;
    timeout?: number;
  },
  callback: (error: Error | null, stdout: string | Buffer, stderr: string | Buffer) => void,
) => { kill: (signal?: NodeJS.Signals) => unknown };

type CancellationTokenLike = {
  createPromise: <T>(
    callback: (
      resolve: (value: T | PromiseLike<T>) => void,
      reject: (error: Error) => void,
      onCancel: (callback: () => void) => void,
    ) => void,
  ) => Promise<T>;
};

type CurlDownloadOptions = {
  headers?: OutgoingHttpHeaders | null;
  sha2?: string | null;
  sha512?: string | null;
  cancellationToken: CancellationTokenLike;
};

export type CurlHttpExecutor = {
  request: (
    options: RequestOptions,
    cancellationToken?: CancellationTokenLike,
    data?: Record<string, unknown> | null,
  ) => Promise<string | null>;
  download: (url: URL, destination: string, options: CurlDownloadOptions) => Promise<string>;
  downloadToBuffer: (url: URL, options: CurlDownloadOptions) => Promise<Buffer>;
};

function requestOptionsToUrl(options: RequestOptions): string {
  const protocol = options.protocol ?? 'https:';
  const hostname = options.hostname ?? options.host;
  if (!hostname) throw new Error('Updater request is missing a hostname.');
  const port = options.port ? `:${options.port}` : '';
  const requestPath = options.path ?? '/';
  return `${protocol}//${hostname}${port}${requestPath}`;
}

function addHeaderArgs(
  args: string[],
  headers: RequestOptions['headers'] | OutgoingHttpHeaders | null | undefined,
): void {
  if (Array.isArray(headers)) {
    for (let index = 0; index < headers.length; index += 2) {
      const name = headers[index];
      const value = headers[index + 1];
      if (name !== undefined && value !== undefined) {
        args.push('--header', `${name}: ${value}`);
      }
    }
    return;
  }
  for (const [name, value] of Object.entries(headers ?? {})) {
    if (value === undefined) continue;
    const values = Array.isArray(value) ? value : [value];
    for (const item of values) {
      args.push('--header', `${name}: ${String(item)}`);
    }
  }
}

function buildBaseArgs(timeoutMs?: number): string[] {
  const args = ['--fail', '--location', '--silent', '--show-error', '--connect-timeout', '30'];
  if (typeof timeoutMs === 'number' && timeoutMs > 0) {
    args.push('--max-time', String(Math.max(1, Math.ceil(timeoutMs / 1000))));
  }
  return args;
}

function runCurl<T>(options: {
  execFile: CurlExecFile;
  curlPath: string;
  args: readonly string[];
  encoding: 'utf8' | 'buffer';
  maxBuffer?: number;
  timeout?: number;
  cancellationToken?: CancellationTokenLike;
}): Promise<T> {
  const run = (
    resolve: (value: T) => void,
    reject: (error: Error) => void,
    onCancel: (callback: () => void) => void,
  ) => {
    const child = options.execFile(
      options.curlPath,
      options.args,
      {
        encoding: options.encoding,
        maxBuffer: options.maxBuffer,
        timeout: options.timeout,
      },
      (error, stdout, stderr) => {
        if (error) {
          const stderrMessage = Buffer.isBuffer(stderr) ? stderr.toString('utf8') : stderr;
          const errno = (error as NodeJS.ErrnoException).code;
          const safeFallback = errno ? `curl failed (${errno})` : 'curl failed';
          reject(new Error(stderrMessage.trim() || safeFallback));
          return;
        }
        resolve(stdout as T);
      },
    );
    onCancel(() => {
      child.kill('SIGTERM');
    });
  };

  if (options.cancellationToken) {
    return options.cancellationToken.createPromise<T>(run);
  }
  return new Promise<T>((resolve, reject) => run(resolve, reject, () => {}));
}

export function createCurlHttpExecutor(
  options: {
    execFile?: CurlExecFile;
    curlPath?: string;
    mkdir?: (targetPath: string) => Promise<unknown>;
    readFile?: (targetPath: string) => Promise<Buffer>;
  } = {},
): CurlHttpExecutor {
  const execFile = options.execFile ?? (defaultExecFile as unknown as CurlExecFile);
  const curlPath = options.curlPath ?? '/usr/bin/curl';
  const mkdir =
    options.mkdir ?? ((targetPath: string) => fs.promises.mkdir(targetPath, { recursive: true }));
  const readFile = options.readFile ?? ((targetPath: string) => fs.promises.readFile(targetPath));

  async function verifyDownloadedFile(destination: string, downloadOptions: CurlDownloadOptions) {
    if (!downloadOptions.sha512 && !downloadOptions.sha2) return;
    const data = await readFile(destination);
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

  return {
    async request(requestOptions, cancellationToken, data): Promise<string | null> {
      const args = buildBaseArgs(requestOptions.timeout);
      addHeaderArgs(args, requestOptions.headers);
      if (requestOptions.method && requestOptions.method !== 'GET') {
        args.push('--request', requestOptions.method);
      }
      if (data) {
        args.push('--data-binary', JSON.stringify(data));
      }
      args.push(requestOptionsToUrl(requestOptions));
      const result = await runCurl<string>({
        execFile,
        curlPath,
        args,
        encoding: 'utf8',
        maxBuffer: 10 * 1024 * 1024,
        timeout: requestOptions.timeout,
        cancellationToken,
      });
      return result.length === 0 ? null : result;
    },
    async download(url, destination, downloadOptions): Promise<string> {
      await mkdir(path.dirname(destination));
      const args = buildBaseArgs();
      addHeaderArgs(args, downloadOptions.headers);
      args.push('--output', destination, url.href);
      await runCurl<Buffer>({
        execFile,
        curlPath,
        args,
        encoding: 'buffer',
        maxBuffer: 1024 * 1024,
        cancellationToken: downloadOptions.cancellationToken,
      });
      await verifyDownloadedFile(destination, downloadOptions);
      return destination;
    },
    async downloadToBuffer(url, downloadOptions): Promise<Buffer> {
      const args = buildBaseArgs();
      addHeaderArgs(args, downloadOptions.headers);
      args.push(url.href);
      return await runCurl<Buffer>({
        execFile,
        curlPath,
        args,
        encoding: 'buffer',
        maxBuffer: 600 * 1024 * 1024,
        cancellationToken: downloadOptions.cancellationToken,
      });
    },
  };
}
