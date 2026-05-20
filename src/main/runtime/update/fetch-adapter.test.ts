import test from 'node:test';
import assert from 'node:assert/strict';
import { createElectronNetFetch, createGlobalFetch } from './fetch-adapter';
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
