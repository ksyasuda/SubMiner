import assert from 'node:assert/strict';
import test from 'node:test';
import { handleCharacterDictionaryAutoSyncComplete } from './character-dictionary-auto-sync-completion';

test('character dictionary sync completion skips expensive subtitle refresh when dictionary is unchanged', () => {
  const calls: string[] = [];

  handleCharacterDictionaryAutoSyncComplete(
    {
      mediaId: 1,
      mediaTitle: 'Frieren',
      changed: false,
    },
    {
      hasParserWindow: () => true,
      clearParserCaches: () => calls.push('clear-parser'),
      invalidateTokenizationCache: () => calls.push('invalidate'),
      refreshSubtitlePrefetch: () => calls.push('prefetch'),
      refreshCurrentSubtitle: () => calls.push('refresh-subtitle'),
      logInfo: (message) => calls.push(`log:${message}`),
    },
  );

  assert.deepEqual(calls, [
    'log:[dictionary:auto-sync] refreshed current subtitle after sync (AniList 1, changed=no, title=Frieren)',
  ]);
});

test('character dictionary sync completion refreshes subtitle state when dictionary changed', () => {
  const calls: string[] = [];

  handleCharacterDictionaryAutoSyncComplete(
    {
      mediaId: 1,
      mediaTitle: 'Frieren',
      changed: true,
    },
    {
      hasParserWindow: () => true,
      clearParserCaches: () => calls.push('clear-parser'),
      invalidateTokenizationCache: () => calls.push('invalidate'),
      refreshSubtitlePrefetch: () => calls.push('prefetch'),
      refreshCurrentSubtitle: () => calls.push('refresh-subtitle'),
      logInfo: (message) => calls.push(`log:${message}`),
    },
  );

  assert.deepEqual(calls, [
    'clear-parser',
    'invalidate',
    'prefetch',
    'refresh-subtitle',
    'log:[dictionary:auto-sync] refreshed current subtitle after sync (AniList 1, changed=yes, title=Frieren)',
  ]);
});
