import test from 'node:test';
import assert from 'node:assert/strict';
import type { AnkiConnectConfig } from '../../types';
import {
  getPreferredYomitanAnkiServerUrl,
  shouldForceOverrideYomitanAnkiServer,
} from './yomitan-anki-server';

function createConfig(overrides: Partial<AnkiConnectConfig> = {}): AnkiConnectConfig {
  return {
    enabled: false,
    url: 'http://127.0.0.1:8765',
    proxy: {
      enabled: true,
      host: '127.0.0.1',
      port: 8766,
      upstreamUrl: 'http://127.0.0.1:8765',
    },
    ...overrides,
  } as AnkiConnectConfig;
}

test('prefers upstream AnkiConnect when SubMiner integration is disabled', () => {
  const config = createConfig({
    enabled: false,
    proxy: {
      enabled: true,
      host: '127.0.0.1',
      port: 8766,
      upstreamUrl: 'http://127.0.0.1:8765',
    },
  });

  assert.equal(getPreferredYomitanAnkiServerUrl(config), 'http://127.0.0.1:8765');
  assert.equal(shouldForceOverrideYomitanAnkiServer(config), false);
});

test('prefers SubMiner proxy when SubMiner integration and proxy are enabled', () => {
  const config = createConfig({
    enabled: true,
    proxy: {
      enabled: true,
      host: '127.0.0.1',
      port: 9988,
      upstreamUrl: 'http://127.0.0.1:8765',
    },
  });

  assert.equal(getPreferredYomitanAnkiServerUrl(config), 'http://127.0.0.1:9988');
  assert.equal(shouldForceOverrideYomitanAnkiServer(config), true);
});

test('falls back to upstream AnkiConnect when proxy transport is disabled', () => {
  const config = createConfig({
    enabled: true,
    proxy: {
      enabled: false,
      host: '127.0.0.1',
      port: 8766,
      upstreamUrl: 'http://127.0.0.1:8765',
    },
  });

  assert.equal(getPreferredYomitanAnkiServerUrl(config), 'http://127.0.0.1:8765');
  assert.equal(shouldForceOverrideYomitanAnkiServer(config), false);
});
