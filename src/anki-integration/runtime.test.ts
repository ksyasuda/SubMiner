import test from 'node:test';
import assert from 'node:assert/strict';

import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import type { AnkiConnectConfig } from '../types';
import { AnkiIntegrationRuntime } from './runtime';

function createRuntime(
  config: Partial<AnkiConnectConfig> = {},
  overrides: Partial<ConstructorParameters<typeof AnkiIntegrationRuntime>[0]> = {},
) {
  const calls: string[] = [];

  const runtime = new AnkiIntegrationRuntime({
    initialConfig: config as AnkiConnectConfig,
    pollingRunner: {
      start: () => calls.push('polling:start'),
      stop: () => calls.push('polling:stop'),
    },
    knownWordCache: {
      startLifecycle: () => calls.push('known:start'),
      stopLifecycle: () => calls.push('known:stop'),
      clearKnownWordCacheState: () => calls.push('known:clear'),
    },
    proxyServerFactory: () => ({
      start: ({ host, port, upstreamUrl }) =>
        calls.push(`proxy:start:${host}:${port}:${upstreamUrl}`),
      stop: () => calls.push('proxy:stop'),
    }),
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
    onConfigChanged: () => undefined,
    ...overrides,
  });

  return { runtime, calls };
}

test('AnkiIntegrationRuntime normalizes url and proxy defaults', () => {
  const { runtime } = createRuntime({
    url: '  http://anki.local:8765  ',
    proxy: {
      enabled: true,
      host: ' 0.0.0.0 ',
      port: 7001,
      upstreamUrl: '   ',
    },
  });

  const normalized = runtime.getConfig();

  assert.equal(normalized.url, 'http://anki.local:8765');
  assert.equal(normalized.proxy?.enabled, true);
  assert.equal(normalized.proxy?.host, '0.0.0.0');
  assert.equal(normalized.proxy?.port, 7001);
  assert.equal(normalized.proxy?.upstreamUrl, 'http://anki.local:8765');
  assert.equal(
    normalized.media?.fallbackDuration,
    DEFAULT_ANKI_CONNECT_CONFIG.media.fallbackDuration,
  );
  assert.equal(
    normalized.media?.syncAnimatedImageToWordAudio,
    DEFAULT_ANKI_CONNECT_CONFIG.media.syncAnimatedImageToWordAudio,
  );
});

test('AnkiIntegrationRuntime starts proxy transport when proxy mode is enabled', () => {
  const { runtime, calls } = createRuntime({
    proxy: {
      enabled: true,
      host: '127.0.0.1',
      port: 9999,
      upstreamUrl: 'http://upstream:8765',
    },
  });

  runtime.start();

  assert.deepEqual(calls, ['known:start', 'proxy:start:127.0.0.1:9999:http://upstream:8765']);
});

test('AnkiIntegrationRuntime switches transports and clears known words when runtime patch disables highlighting', () => {
  const { runtime, calls } = createRuntime({
    knownWords: {
      highlightEnabled: true,
    },
    pollingRate: 250,
  });

  runtime.start();
  calls.length = 0;

  runtime.applyRuntimeConfigPatch({
    knownWords: {
      highlightEnabled: false,
    },
    proxy: {
      enabled: true,
      host: '127.0.0.1',
      port: 8766,
      upstreamUrl: 'http://127.0.0.1:8765',
    },
  });

  assert.deepEqual(calls, [
    'known:stop',
    'known:clear',
    'polling:stop',
    'proxy:start:127.0.0.1:8766:http://127.0.0.1:8765',
  ]);
});

test('AnkiIntegrationRuntime skips known-word lifecycle restart for unrelated runtime patches', () => {
  const { runtime, calls } = createRuntime({
    knownWords: {
      highlightEnabled: true,
    },
    pollingRate: 250,
  });

  runtime.start();
  calls.length = 0;

  runtime.applyRuntimeConfigPatch({
    behavior: {
      autoUpdateNewCards: false,
    },
  });

  assert.deepEqual(calls, []);
});

test('AnkiIntegrationRuntime restarts known-word lifecycle when known-word settings change', () => {
  const { runtime, calls } = createRuntime({
    knownWords: {
      highlightEnabled: true,
      refreshMinutes: 90,
    },
    pollingRate: 250,
  });

  runtime.start();
  calls.length = 0;

  runtime.applyRuntimeConfigPatch({
    knownWords: {
      refreshMinutes: 120,
    },
  });

  assert.deepEqual(calls, ['known:start']);
});

test('AnkiIntegrationRuntime does not stop lifecycle when disabled while runtime is stopped', () => {
  const { runtime, calls } = createRuntime({
    knownWords: {
      highlightEnabled: true,
    },
  });

  runtime.applyRuntimeConfigPatch({
    knownWords: {
      highlightEnabled: false,
    },
  });

  assert.deepEqual(calls, ['known:clear']);
});

test('AnkiIntegrationRuntime does not restart known-word lifecycle for config changes while stopped', () => {
  const { runtime, calls } = createRuntime({
    knownWords: {
      highlightEnabled: true,
      refreshMinutes: 90,
    },
  });

  runtime.applyRuntimeConfigPatch({
    knownWords: {
      refreshMinutes: 120,
    },
  });

  assert.deepEqual(calls, []);
});
