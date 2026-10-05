import test from 'node:test';
import assert from 'node:assert/strict';
import { createMineSentenceCardHandler, createRefreshKnownWordCacheHandler } from './anki-actions';

test('refresh known word cache handler throws when Anki integration missing', async () => {
  const refresh = createRefreshKnownWordCacheHandler({
    getAnkiIntegration: () => null,
    missingIntegrationMessage: 'AnkiConnect integration not enabled',
  });

  await assert.rejects(() => refresh(), /AnkiConnect integration not enabled/);
});

test('mine sentence handler forwards the canonical primary subtitle snapshot', async () => {
  const primarySubtitle = { text: '正式な字幕', startTime: 1, endTime: 3 };
  const mineSentenceCard = createMineSentenceCardHandler({
    getAnkiIntegration: () => ({}),
    getMpvClient: () => ({}),
    getPrimarySubtitle: () => primarySubtitle,
    showMpvOsd: () => {},
    mineSentenceCardCore: async (options) => {
      assert.equal(options.primarySubtitle, primarySubtitle);
      return true;
    },
  });

  await mineSentenceCard();
});
