import assert from 'node:assert/strict';
import test from 'node:test';
import { buildYomitanAnkiSettingsKey } from './yomitan-anki-server-sync';
import { buildHachidoriAnkiHints } from '../../core/services/tokenizer/hachidori-anki-settings';

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

test('settings sync key changes when fields or tags change', () => {
  const key = (word: string, tags: string[]) =>
    buildYomitanAnkiSettingsKey({
      targetUrl: 'http://127.0.0.1:8766',
      targetDeck: 'Mining',
      forceOverride: true,
      hachidoriHints: buildHachidoriAnkiHints({ fields: { word }, tags }),
    });
  assert.notEqual(key('Word', ['SubMiner']), key('Expression', ['SubMiner']));
  assert.notEqual(key('Word', ['SubMiner']), key('Word', ['Japanese']));
});
