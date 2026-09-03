import test from 'node:test';
import assert from 'node:assert/strict';
import { AnimeBridgeClient, BridgeExtensionError } from './bridge-client';
import { BRIDGE_CONTEXT_KEY } from './types';

const EXTENSION_ID = 'a'.repeat(64);
const APK_BASE64 = 'QVBLLUJZVEVT';
const source = {
  fingerprint: 'sha-1',
  loadApkBase64: async () => APK_BASE64,
  sourceId: 'source-1',
};

interface Recorded {
  url: string;
  body: Record<string, unknown>;
}

function stubFetch(responder: (call: Recorded, index: number) => Response): {
  fetchImpl: typeof fetch;
  calls: Recorded[];
} {
  const calls: Recorded[] = [];
  const fetchImpl = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const call: Recorded = {
      url: String(input),
      body: init?.body ? (JSON.parse(String(init.body)) as Record<string, unknown>) : {},
    };
    calls.push(call);
    return responder(call, calls.length - 1);
  }) as typeof fetch;
  return { fetchImpl, calls };
}

function jsonResponse(body: unknown, extensionId?: string): Response {
  const headers = new Headers({ 'Content-Type': 'application/json' });
  if (extensionId) headers.set('x-mangatan-extension-id', extensionId);
  return new Response(JSON.stringify(body), { status: 200, headers });
}

test('isReady requires every capability the client depends on', async () => {
  const ready = new AnimeBridgeClient({
    baseUrl: 'http://127.0.0.1:9',
    fetchImpl: stubFetch(() =>
      jsonResponse({ mangatanMihonBridge: 1, sourceFactory: true, preferenceCallbacks: true }),
    ).fetchImpl,
  });
  assert.equal(await ready.isReady(), true);

  const partial = new AnimeBridgeClient({
    baseUrl: 'http://127.0.0.1:9',
    fetchImpl: stubFetch(() => jsonResponse({ mangatanMihonBridge: 1, sourceFactory: true }))
      .fetchImpl,
  });
  assert.equal(await partial.isReady(), false);
});

test('isReady caps the probe at the deadline the caller passes', async () => {
  // A bridge that accepts the socket and then stalls: only the abort ends it.
  const fetchImpl = (async (_input: RequestInfo | URL, init?: RequestInit) => {
    await new Promise((resolve) => init?.signal?.addEventListener('abort', resolve));
    throw new Error('aborted');
  }) as typeof fetch;
  const client = new AnimeBridgeClient({ baseUrl: 'http://127.0.0.1:9', fetchImpl });

  const started = Date.now();
  assert.equal(await client.isReady(50), false);
  // Well under the 5s default, so the per-call deadline is what applied.
  assert.ok(Date.now() - started < 1000, 'probe outlived the caller deadline');
});

test('isReady reports false instead of throwing when the bridge is down', async () => {
  const client = new AnimeBridgeClient({
    baseUrl: 'http://127.0.0.1:9',
    fetchImpl: (async () => {
      throw new Error('ECONNREFUSED');
    }) as typeof fetch,
  });
  assert.equal(await client.isReady(), false);
});

test('getVideoList posts the APK and episode url with a bridge context preference', async () => {
  const { fetchImpl, calls } = stubFetch(() => jsonResponse([{ videoUrl: 'http://x/video/t' }]));
  const client = new AnimeBridgeClient({ baseUrl: 'http://127.0.0.1:9/', fetchImpl });

  const videos = await client.getVideoList(source, 'https://origin.example/ep/1');

  assert.equal(calls[0]?.url, 'http://127.0.0.1:9/dalvik');
  assert.equal(calls[0]?.body.method, 'getVideoList');
  assert.deepEqual(calls[0]?.body.episodeData, { url: 'https://origin.example/ep/1' });
  assert.equal(calls[0]?.body.data, APK_BASE64);
  assert.deepEqual(calls[0]?.body.preferences, [{ key: BRIDGE_CONTEXT_KEY, sourceId: 'source-1' }]);
  assert.equal(videos.length, 1);
});

test('a cached extension id replaces the APK upload on later calls', async () => {
  const { fetchImpl, calls } = stubFetch(() => jsonResponse([], EXTENSION_ID));
  const client = new AnimeBridgeClient({ baseUrl: 'http://127.0.0.1:9', fetchImpl });

  await client.getVideoList(source, 'https://origin.example/ep/1');
  await client.getVideoList(source, 'https://origin.example/ep/2');

  assert.equal(calls[0]?.body.data, APK_BASE64);
  assert.equal(calls[0]?.body.extensionId, undefined);
  assert.equal(calls[1]?.body.data, undefined);
  assert.equal(calls[1]?.body.extensionId, EXTENSION_ID);
});

test('an upgraded APK re-uploads instead of reusing the previous extension id', async () => {
  const { fetchImpl, calls } = stubFetch(() => jsonResponse([], EXTENSION_ID));
  const client = new AnimeBridgeClient({ baseUrl: 'http://127.0.0.1:9', fetchImpl });

  await client.getVideoList(source, 'https://origin.example/ep/1');
  // Same source id, new build in the same file: the id cache must miss.
  const upgraded = { ...source, fingerprint: 'sha-2', loadApkBase64: async () => 'TkVXLUFQSw==' };
  await client.getVideoList(upgraded, 'https://origin.example/ep/2');

  assert.equal(calls[1]?.body.extensionId, undefined);
  assert.equal(calls[1]?.body.data, 'TkVXLUFQSw==');
});

test('a 409 re-uploads the APK once and succeeds', async () => {
  const { fetchImpl, calls } = stubFetch((call, index) => {
    if (index === 0) return jsonResponse([], EXTENSION_ID);
    // Cache evicted: reject the id-only call, accept the re-upload.
    if (call.body.extensionId !== undefined) return new Response('', { status: 409 });
    return jsonResponse([{ videoUrl: 'http://x/video/t' }], EXTENSION_ID);
  });
  const client = new AnimeBridgeClient({ baseUrl: 'http://127.0.0.1:9', fetchImpl });

  await client.getVideoList(source, 'https://origin.example/ep/1');
  const videos = await client.getVideoList(source, 'https://origin.example/ep/2');

  assert.equal(calls.length, 3);
  assert.equal(calls[1]?.body.extensionId, EXTENSION_ID);
  assert.equal(calls[2]?.body.data, APK_BASE64);
  assert.equal(videos.length, 1);
});

test('an error body on a 200 response raises BridgeExtensionError with the code', async () => {
  const { fetchImpl } = stubFetch(() => jsonResponse({ error: 'Cloudflare challenge', code: 403 }));
  const client = new AnimeBridgeClient({ baseUrl: 'http://127.0.0.1:9', fetchImpl });

  await assert.rejects(
    () => client.getVideoList(source, 'https://origin.example/ep/1'),
    (error: unknown) => {
      assert.ok(error instanceof BridgeExtensionError);
      assert.equal(error.code, 403);
      assert.match(error.message, /Cloudflare challenge/);
      return true;
    },
  );
});

test('searchAnime sends a 1-based page and returns the page payload', async () => {
  const { fetchImpl, calls } = stubFetch(() =>
    jsonResponse({ animes: [{ title: 'Example' }], hasNextPage: true }),
  );
  const client = new AnimeBridgeClient({ baseUrl: 'http://127.0.0.1:9', fetchImpl });

  const page = await client.searchAnime(source, 'example');

  assert.equal(calls[0]?.body.method, 'getSearchAnime');
  assert.equal(calls[0]?.body.page, 1);
  assert.equal(calls[0]?.body.search, 'example');
  assert.deepEqual(calls[0]?.body.filterList, []);
  assert.equal(page.hasNextPage, true);
  assert.equal(page.animes?.length, 1);
});

test('getEpisodeList wraps the anime url in animeData', async () => {
  const { fetchImpl, calls } = stubFetch(() => jsonResponse([{ name: 'Episode 1', url: '/ep/1' }]));
  const client = new AnimeBridgeClient({ baseUrl: 'http://127.0.0.1:9', fetchImpl });

  const episodes = await client.getEpisodeList(source, 'https://origin.example/anime/1');

  assert.equal(calls[0]?.body.method, 'getEpisodeList');
  assert.deepEqual(calls[0]?.body.animeData, { url: 'https://origin.example/anime/1' });
  assert.equal(episodes[0]?.name, 'Episode 1');
});
