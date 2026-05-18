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
