import assert from 'node:assert/strict';
import test from 'node:test';
import type { AnkiConnectConfig, ResolvedConfig } from '../../types';
import { createYomitanAnkiServerSyncRuntime } from './yomitan-anki-server-sync';

type SyncCall = { url: string; forceOverride: boolean; deck: string };

function createHarness(options: { readOnly?: boolean; syncResults?: boolean[] } = {}) {
  let ankiConnect = { url: 'http://127.0.0.1:8765', deck: 'Mining' } as AnkiConnectConfig;
  const syncResults = [...(options.syncResults ?? [])];
  const calls: SyncCall[] = [];
  const runtime = createYomitanAnkiServerSyncRuntime({
    isExternalReadOnlyMode: () => options.readOnly ?? false,
    getResolvedConfig: () => ({ ankiConnect }) as ResolvedConfig,
    getYomitanParserRuntimeDeps: () => ({}) as never,
    logError: () => {},
    logInfo: () => {},
    syncDefaultAnkiServer: async (url, _deps, _logger, syncOptions) => {
      calls.push({
        url,
        forceOverride: syncOptions?.forceOverride ?? false,
        deck: syncOptions?.deck ?? '',
      });
      return syncResults.shift() ?? true;
    },
  });
  return {
    calls,
    sync: () => runtime.syncYomitanDefaultProfileAnkiServer(),
    setAnkiConnect: (next: Partial<AnkiConnectConfig>) => {
      ankiConnect = { ...ankiConnect, ...next } as AnkiConnectConfig;
    },
  };
}

test('skips syncing in external read-only mode', async () => {
  const harness = createHarness({ readOnly: true });
  await harness.sync();
  assert.deepEqual(harness.calls, []);
});

test('syncs once per distinct url, deck, and override settings', async () => {
  const harness = createHarness();
  await harness.sync();
  await harness.sync();
  assert.deepEqual(harness.calls, [
    { url: 'http://127.0.0.1:8765', forceOverride: false, deck: 'Mining' },
  ]);

  harness.setAnkiConnect({ deck: ' Sentences ' });
  await harness.sync();
  harness.setAnkiConnect({ enabled: true, proxy: { enabled: true, host: '', port: 0 } } as never);
  await harness.sync();
  assert.deepEqual(harness.calls.slice(1), [
    { url: 'http://127.0.0.1:8765', forceOverride: false, deck: 'Sentences' },
    { url: 'http://127.0.0.1:8766', forceOverride: true, deck: 'Sentences' },
  ]);
});

test('retries the same settings after a failed sync', async () => {
  const harness = createHarness({ syncResults: [false, true] });
  await harness.sync();
  await harness.sync();
  await harness.sync();
  assert.equal(harness.calls.length, 2);
});
