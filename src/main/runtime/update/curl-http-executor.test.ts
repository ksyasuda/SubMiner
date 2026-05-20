import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { createCurlHttpExecutor, type CurlExecFile } from './curl-http-executor';

test('curl HTTP executor requests updater metadata without Electron networking', async () => {
  const calls: Array<{ file: string; args: readonly string[] }> = [];
  const execFile: CurlExecFile = (file, args, _options, callback) => {
    calls.push({ file, args });
    queueMicrotask(() => callback(null, 'metadata', ''));
    return { kill: () => true };
  };
  const executor = createCurlHttpExecutor({ execFile, curlPath: '/usr/bin/curl' });

  const result = await executor.request({
    protocol: 'https:',
    hostname: 'api.github.com',
    path: '/repos/ksyasuda/SubMiner/releases',
    headers: {
      Accept: 'application/vnd.github+json',
      'x-user-staging-id': 'abc',
    },
    timeout: 120_000,
  });

  assert.equal(result, 'metadata');
  assert.equal(calls.length, 1);
  assert.equal(calls[0]?.file, '/usr/bin/curl');
  assert.deepEqual(calls[0]?.args, [
    '--fail',
    '--location',
    '--silent',
    '--show-error',
    '--connect-timeout',
    '30',
    '--max-time',
    '120',
    '--header',
    'Accept: application/vnd.github+json',
    '--header',
    'x-user-staging-id: abc',
    'https://api.github.com/repos/ksyasuda/SubMiner/releases',
  ]);
});

test('curl HTTP executor downloads updater assets to the requested destination', async () => {
  const calls: Array<{ args: readonly string[]; timeout?: number }> = [];
  const execFile: CurlExecFile = (_file, args, options, callback) => {
    calls.push({ args, timeout: options.timeout });
    queueMicrotask(() => callback(null, Buffer.alloc(0), Buffer.alloc(0)));
    return { kill: () => true };
  };
  const executor = createCurlHttpExecutor({
    execFile,
    curlPath: '/usr/bin/curl',
    mkdir: async () => undefined,
    downloadTimeoutMs: 120_000,
  });

  await executor.download(
    new URL('https://github.com/ksyasuda/SubMiner/releases/download/v1/app.zip'),
    '/tmp/subminer/update.zip',
    {
      headers: { 'User-Agent': 'SubMiner updater' },
      cancellationToken: {
        createPromise: (callback) =>
          new Promise((resolve, reject) => callback(resolve, reject, () => {})),
      },
    },
  );

  assert.deepEqual(calls[0]?.args, [
    '--fail',
    '--location',
    '--silent',
    '--show-error',
    '--connect-timeout',
    '30',
    '--max-time',
    '120',
    '--header',
    'User-Agent: SubMiner updater',
    '--output',
    '/tmp/subminer/update.zip',
    'https://github.com/ksyasuda/SubMiner/releases/download/v1/app.zip',
  ]);
  assert.equal(calls[0]?.timeout, 120_000);
});

test('curl HTTP executor verifies downloaded updater asset hashes', async () => {
  const data = Buffer.from('zip payload');
  const expectedSha512 = createHash('sha512').update(data).digest('base64');
  const execFile: CurlExecFile = (_file, _args, _options, callback) => {
    queueMicrotask(() => callback(null, Buffer.alloc(0), Buffer.alloc(0)));
    return { kill: () => true };
  };
  const executor = createCurlHttpExecutor({
    execFile,
    curlPath: '/usr/bin/curl',
    mkdir: async () => undefined,
    readFile: async () => data,
  });

  await executor.download(new URL('https://example.test/update.zip'), '/tmp/subminer/update.zip', {
    sha512: expectedSha512,
    cancellationToken: {
      createPromise: (callback) =>
        new Promise((resolve, reject) => callback(resolve, reject, () => {})),
    },
  });

  await assert.rejects(
    () =>
      executor.download(new URL('https://example.test/update.zip'), '/tmp/subminer/update.zip', {
        sha512: 'bad',
        cancellationToken: {
          createPromise: (callback) =>
            new Promise((resolve, reject) => callback(resolve, reject, () => {})),
        },
      }),
    /sha512 mismatch/,
  );
});

test('curl HTTP executor does not expose command arguments when stderr is empty', async () => {
  const execFile: CurlExecFile = (_file, _args, _options, callback) => {
    const error = new Error('--header Authorization: Bearer secret-token');
    Object.assign(error, { code: 'ENOENT' });
    queueMicrotask(() => callback(error, '', ''));
    return { kill: () => true };
  };
  const executor = createCurlHttpExecutor({ execFile, curlPath: '/usr/bin/curl' });

  await assert.rejects(
    () =>
      executor.request({
        protocol: 'https:',
        hostname: 'api.github.com',
        path: '/repos/ksyasuda/SubMiner/releases',
      }),
    (error) => {
      assert.ok(error instanceof Error);
      assert.equal(error.message, 'curl failed (ENOENT)');
      assert.doesNotMatch(error.message, /secret-token|Authorization/);
      return true;
    },
  );
});
