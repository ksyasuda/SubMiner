import assert from 'node:assert/strict';
import test from 'node:test';
import { buildYomitanAnkiSettingsKey } from './yomitan-anki-server-sync';

test('buildYomitanAnkiSettingsKey includes force override policy', () => {
  assert.notEqual(
    buildYomitanAnkiSettingsKey({
      targetUrl: 'http://127.0.0.1:8766',
      targetDeck: 'Mining',
      forceOverride: false,
    }),
    buildYomitanAnkiSettingsKey({
      targetUrl: 'http://127.0.0.1:8766',
      targetDeck: 'Mining',
      forceOverride: true,
    }),
  );
});
