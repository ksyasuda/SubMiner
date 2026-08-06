import test from 'node:test';
import assert from 'node:assert/strict';
import { createCurlFetch, createElectronNetFetch, createGlobalFetch } from './fetch-adapter';
import type { FetchResponseLike } from './release-assets';

test('createElectronNetFetch delegates updater requests to Electron net.fetch', async () => {
  const calls: Array<{ url: string; init?: Record<string, unknown> }> = [];
  const response: FetchResponseLike = {
    ok: true,
    status: 200,
    statusText: 'OK',
    json: async () => ({ ok: true }),
    text: async () => 'ok',
    arrayBuffer: async () => new ArrayBuffer(0),
  };

  const fetch = createElectronNetFetch({
    fetch: async (url, init) => {
      calls.push({ url, init });
      return response;
    },
  });

  const result = await fetch('https://api.github.com/repos/ksyasuda/SubMiner/releases', {
    headers: { 'User-Agent': 'SubMiner updater' },
  });

  assert.equal(result, response);
  assert.deepEqual(calls, [
    {
      url: 'https://api.github.com/repos/ksyasuda/SubMiner/releases',
      init: { headers: { 'User-Agent': 'SubMiner updater' } },
    },
  ]);
});

test('createGlobalFetch delegates updater requests to main-process fetch', async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = [];
  const response: FetchResponseLike = {
    ok: true,
    status: 200,
    statusText: 'OK',
    json: async () => ({ ok: true }),
    text: async () => 'ok',
    arrayBuffer: async () => new ArrayBuffer(0),
  };

  const fetch = createGlobalFetch(async (url, init) => {
    calls.push({ url, init });
    return response;
  });

  const result = await fetch('https://api.github.com/repos/ksyasuda/SubMiner/releases', {
    headers: { 'User-Agent': 'SubMiner updater' },
  });

  assert.equal(result, response);
  assert.deepEqual(calls, [
    {
      url: 'https://api.github.com/repos/ksyasuda/SubMiner/releases',
      init: { headers: { 'User-Agent': 'SubMiner updater' } } as RequestInit,
    },
  ]);
});

test('createCurlFetch requests updater metadata without Electron networking', async () => {
  const calls: Array<{
    file: string;
    args: readonly string[];
    options: { encoding: 'utf8' | 'buffer'; maxBuffer?: number; timeout?: number };
  }> = [];
  const payload = Buffer.from(JSON.stringify([{ tag_name: 'v1.2.3', assets: [] }]));

  const fetch = createCurlFetch({
    curlPath: '/usr/bin/curl',
    execFile: (file, args, options, callback) => {
      calls.push({ file, args, options });
      callback(null, payload, Buffer.alloc(0));
      return { kill: () => undefined };
    },
  });

  const response = await fetch('https://api.github.com/repos/ksyasuda/SubMiner/releases', {
    headers: {
      Accept: 'application/vnd.github+json',
      'User-Agent': 'SubMiner updater',
    },
  });

  assert.equal(response.ok, true);
  assert.equal(response.status, 200);
  assert.deepEqual(await response.json(), [{ tag_name: 'v1.2.3', assets: [] }]);
  assert.equal(await response.text(), '[{"tag_name":"v1.2.3","assets":[]}]');
  assert.deepEqual(Buffer.from(await response.arrayBuffer()), payload);
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
    '60',
    '--header',
    'Accept: application/vnd.github+json',
    '--header',
    'User-Agent: SubMiner updater',
    'https://api.github.com/repos/ksyasuda/SubMiner/releases',
  ]);
  assert.equal(calls[0]?.options.encoding, 'buffer');
  assert.equal(calls[0]?.options.timeout, 65_000);
});

test('curl fetch kills the child process when the caller aborts', async () => {
  let killed: string | undefined;
  let settle: ((error: Error | null, stdout: Buffer, stderr: Buffer) => void) | null = null;
  const controller = new AbortController();

  const curlFetch = createCurlFetch({
    execFile: ((
      _file: string,
      _args: readonly string[],
      _options: unknown,
      callback: (error: Error | null, stdout: Buffer, stderr: Buffer) => void,
    ) => {
      settle = callback;
      return {
        kill: (signal?: string) => {
          killed = signal;
          return true;
        },
      };
    }) as never,
  });

  const pending = curlFetch('https://example.test/slow', { signal: controller.signal });
  controller.abort(new Error('deadline reached'));

  await assert.rejects(pending, /deadline reached/);
  assert.equal(killed, 'SIGKILL', 'the stalled curl process is terminated');
  assert.ok(settle, 'the callback is still held by the stub');
});

test('curl fetch preserves a non-Error abort reason', async () => {
  const controller = new AbortController();
  const curlFetch = createCurlFetch({
    execFile: (() => ({ kill: () => true })) as never,
  });

  const pending = curlFetch('https://example.test/slow', { signal: controller.signal });
  controller.abort('plain string reason');

  await pending.then(
    () => assert.fail('expected the aborted request to reject'),
    (error) => assert.equal(error, 'plain string reason'),
  );
});

test('curl fetch synthesises an error when the signal carries no reason', async () => {
  const controller = new AbortController();
  const curlFetch = createCurlFetch({
    execFile: (() => ({ kill: () => true })) as never,
  });

  const pending = curlFetch('https://example.test/slow', { signal: controller.signal });
  controller.abort();

  // A bare abort() still supplies a DOMException reason, so that is what surfaces.
  await pending.then(
    () => assert.fail('expected the aborted request to reject'),
    (error) => assert.ok(error instanceof Error),
  );
});

test('curl fetch rejects immediately when the signal is already aborted', async () => {
  let spawned = 0;
  const controller = new AbortController();
  controller.abort(new Error('already gone'));

  const curlFetch = createCurlFetch({
    execFile: (() => {
      spawned += 1;
      return { kill: () => true };
    }) as never,
  });

  await assert.rejects(
    curlFetch('https://example.test/x', { signal: controller.signal }),
    /already gone/,
  );
  assert.equal(spawned, 0, 'no curl process is started for an aborted request');
});

test('curl fetch still resolves normally when no signal is supplied', async () => {
  const curlFetch = createCurlFetch({
    execFile: ((
      _file: string,
      _args: readonly string[],
      _options: unknown,
      callback: (error: Error | null, stdout: Buffer, stderr: Buffer) => void,
    ) => {
      callback(null, Buffer.from('{"ok":true}'), Buffer.alloc(0));
      return { kill: () => true };
    }) as never,
  });

  const response = await curlFetch('https://example.test/x');
  assert.equal(response.ok, true);
  assert.deepEqual(await response.json(), { ok: true });
});
