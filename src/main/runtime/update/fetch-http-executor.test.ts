import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import test from 'node:test';
import { createFetchHttpExecutor } from './fetch-http-executor';

function neverCancel<T>(
  callback: (
    resolve: (value: T | PromiseLike<T>) => void,
    reject: (error: Error) => void,
    onCancel: (callback: () => void) => void,
  ) => void,
): Promise<T> {
  return new Promise<T>((resolve, reject) => callback(resolve, reject, () => {}));
}

test('fetch HTTP executor requests updater metadata without Electron networking', async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = [];
  const executor = createFetchHttpExecutor({
    fetch: async (url, init) => {
      calls.push({ url, init });
      return new Response('version: 0.15.0');
    },
  });

  const result = await executor.request({
    protocol: 'https:',
    hostname: 'example.test',
    path: '/latest.yml',
    headers: { Accept: 'application/octet-stream' },
    timeout: 5_000,
  });

  assert.equal(result, 'version: 0.15.0');
  assert.equal(calls[0]?.url, 'https://example.test/latest.yml');
  assert.equal((calls[0]?.init?.headers as Headers).get('Accept'), 'application/octet-stream');
  assert.equal(calls[0]?.init?.method, 'GET');
});

test('fetch HTTP executor downloads updater assets to the requested destination', async () => {
  const data = Buffer.from('installer');
  const written: Array<{ path: string; data: Buffer }> = [];
  const executor = createFetchHttpExecutor({
    fetch: async () => new Response(new Uint8Array(data)),
    mkdir: async () => undefined,
    writeFile: async (targetPath, body) => {
      written.push({ path: targetPath, data: body });
    },
  });

  const destination = await executor.download(
    new URL('https://example.test/SubMiner-0.15.0.exe'),
    'C:\\Temp\\SubMiner-0.15.0.exe',
    {
      cancellationToken: {
        createPromise: neverCancel,
      },
      sha2: createHash('sha256').update(data).digest('hex'),
    },
  );

  assert.equal(destination, 'C:\\Temp\\SubMiner-0.15.0.exe');
  assert.deepEqual(written, [{ path: destination, data }]);
});

test('fetch HTTP executor verifies updater asset hashes', async () => {
  const executor = createFetchHttpExecutor({
    fetch: async () => new Response('wrong data'),
    mkdir: async () => undefined,
    writeFile: async () => {
      throw new Error('should not write mismatched data');
    },
  });

  await assert.rejects(
    executor.download(new URL('https://example.test/SubMiner.exe'), 'C:\\Temp\\SubMiner.exe', {
      cancellationToken: {
        createPromise: neverCancel,
      },
      sha2: createHash('sha256').update('expected data').digest('hex'),
    }),
    /sha2 mismatch/,
  );
});

test('fetch HTTP executor applies download timeout to updater asset fetches', async () => {
  const executor = createFetchHttpExecutor({
    downloadTimeoutMs: 1,
    fetch: async (_url, init) =>
      new Promise<Response>((_resolve, reject) => {
        init?.signal?.addEventListener('abort', () => reject(new Error('download aborted')), {
          once: true,
        });
      }),
    mkdir: async () => undefined,
    writeFile: async () => {
      throw new Error('should not write timed-out data');
    },
  });

  await assert.rejects(
    executor.download(new URL('https://example.test/SubMiner.exe'), 'C:\\Temp\\SubMiner.exe', {
      cancellationToken: {
        createPromise: neverCancel,
      },
    }),
    /download aborted/,
  );
});

test('fetch HTTP executor applies download timeout to buffer fetches', async () => {
  const executor = createFetchHttpExecutor({
    downloadTimeoutMs: 1,
    fetch: async (_url, init) =>
      new Promise<Response>((_resolve, reject) => {
        init?.signal?.addEventListener('abort', () => reject(new Error('buffer aborted')), {
          once: true,
        });
      }),
  });

  await assert.rejects(
    executor.downloadToBuffer(new URL('https://example.test/SubMiner.exe'), {
      cancellationToken: {
        createPromise: neverCancel,
      },
    }),
    /buffer aborted/,
  );
});
