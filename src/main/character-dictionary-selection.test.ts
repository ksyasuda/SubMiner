import assert from 'node:assert/strict';
import test from 'node:test';

import { applyCharacterDictionarySelection } from './character-dictionary-selection';

test('applyCharacterDictionarySelection returns saved override when post-save sync fails', async () => {
  const warnings: unknown[] = [];
  const result = await applyCharacterDictionarySelection(
    { mediaId: 21355 },
    {
      setManualSelection: async (request) => ({
        ok: true,
        seriesKey: `series-${request.mediaId}`,
        selected: { id: request.mediaId, title: 'Re:ZERO', episodes: 25 },
        staleMediaIds: [10607],
      }),
      resetAnilistMediaGuessState: () => {},
      runSyncNow: async () => {
        throw new Error('sync failed');
      },
      warn: (...args) => warnings.push(args),
    },
  );

  assert.equal(result.selected.id, 21355);
  assert.equal(warnings.length, 1);
});
