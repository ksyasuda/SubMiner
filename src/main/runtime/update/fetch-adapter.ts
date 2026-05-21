import { execFile as defaultExecFile } from 'node:child_process';
import type { FetchLike, FetchResponseLike } from './release-assets';
import type { CurlExecFile } from './curl-http-executor';

export interface ElectronNetFetchLike {
  fetch: (url: string, init?: Record<string, unknown>) => Promise<FetchResponseLike>;
}

export type GlobalFetchLike = (url: string, init?: RequestInit) => Promise<FetchResponseLike>;

export function createElectronNetFetch(net: ElectronNetFetchLike): FetchLike {
  return (url, init) => net.fetch(url, init);
}

function getGlobalFetch(): GlobalFetchLike {
  if (typeof globalThis.fetch !== 'function') {
    throw new Error('Global fetch is not available for updater requests.');
  }
  return globalThis.fetch.bind(globalThis) as GlobalFetchLike;
}

export function createGlobalFetch(fetchImpl?: GlobalFetchLike): FetchLike {
  return (url, init) => (fetchImpl ?? getGlobalFetch())(url, init as RequestInit);
}

type CurlFetchOptions = {
  execFile?: CurlExecFile;
  curlPath?: string;
};

function addHeaderArgs(args: string[], headers: unknown): void {
  if (!headers) return;
  if (Array.isArray(headers)) {
    for (const header of headers) {
      if (Array.isArray(header) && header.length >= 2) {
        args.push('--header', `${header[0]}: ${header[1]}`);
      }
    }
    return;
  }
  if (typeof headers === 'object' && 'forEach' in headers) {
    (headers as { forEach: (callback: (value: string, name: string) => void) => void }).forEach(
      (value, name) => {
        args.push('--header', `${name}: ${value}`);
      },
    );
    return;
  }
  if (typeof headers !== 'object') return;
  for (const [name, value] of Object.entries(headers as Record<string, unknown>)) {
    if (value === undefined) continue;
    const values = Array.isArray(value) ? value : [value];
    for (const item of values) {
      args.push('--header', `${name}: ${String(item)}`);
    }
  }
}

function bufferToArrayBuffer(buffer: Buffer): ArrayBuffer {
  const result = new ArrayBuffer(buffer.length);
  new Uint8Array(result).set(buffer);
  return result;
}

export function createCurlFetch(options: CurlFetchOptions = {}): FetchLike {
  const execFile = options.execFile ?? (defaultExecFile as unknown as CurlExecFile);
  const curlPath = options.curlPath ?? '/usr/bin/curl';

  return async (url, init = {}) => {
    const args = [
      '--fail',
      '--location',
      '--silent',
      '--show-error',
      '--connect-timeout',
      '30',
      '--max-time',
      '60',
    ];
    addHeaderArgs(args, init.headers);
    args.push(url);
    const body = await new Promise<Buffer>((resolve, reject) => {
      execFile(
        curlPath,
        args,
        {
          encoding: 'buffer',
          maxBuffer: 600 * 1024 * 1024,
          timeout: 65_000,
        },
        (error, stdout, stderr) => {
          if (error) {
            const stderrMessage = Buffer.isBuffer(stderr) ? stderr.toString('utf8') : stderr;
            const errno = (error as NodeJS.ErrnoException).code;
            const fallback = errno ? `curl failed (${errno})` : 'curl failed';
            reject(new Error(stderrMessage.trim() || fallback));
            return;
          }
          resolve(Buffer.isBuffer(stdout) ? stdout : Buffer.from(stdout));
        },
      );
    });
    return {
      ok: true,
      status: 200,
      statusText: 'OK',
      json: async () => JSON.parse(body.toString('utf8')),
      text: async () => body.toString('utf8'),
      arrayBuffer: async () => bufferToArrayBuffer(body),
    };
  };
}
