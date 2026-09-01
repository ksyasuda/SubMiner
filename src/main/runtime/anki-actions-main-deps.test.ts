import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildMarkLastCardAsAudioCardMainDepsHandler,
  createBuildMineSentenceCardMainDepsHandler,
  createBuildRefreshKnownWordCacheMainDepsHandler,
  createBuildTriggerFieldGroupingMainDepsHandler,
  createBuildUpdateLastCardFromClipboardMainDepsHandler,
} from './anki-actions-main-deps';

test('anki action main deps builders map callbacks', async () => {
  const calls: string[] = [];

  const update = createBuildUpdateLastCardFromClipboardMainDepsHandler({
    getAnkiIntegration: () => ({ enabled: true }),
    readClipboardText: () => 'clip',
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    updateLastCardFromClipboardCore: async () => {
      calls.push('update');
    },
  })();
  assert.deepEqual(update.getAnkiIntegration(), { enabled: true });
  assert.equal(update.readClipboardText(), 'clip');
  update.showMpvOsd('x');
  await update.updateLastCardFromClipboardCore({
    ankiIntegration: { enabled: true },
    readClipboardText: () => '',
    showMpvOsd: () => {},
  });

  const refresh = createBuildRefreshKnownWordCacheMainDepsHandler({
    getAnkiIntegration: () => null,
    missingIntegrationMessage: 'missing',
  })();
  assert.equal(refresh.getAnkiIntegration(), null);
  assert.equal(refresh.missingIntegrationMessage, 'missing');

  const fieldGrouping = createBuildTriggerFieldGroupingMainDepsHandler({
    getAnkiIntegration: () => ({ enabled: true }),
    showMpvOsd: (text) => calls.push(`fg:${text}`),
    triggerFieldGroupingCore: async () => {
      calls.push('trigger');
    },
  })();
  fieldGrouping.showMpvOsd('fg');
  await fieldGrouping.triggerFieldGroupingCore({
    ankiIntegration: { enabled: true },
    showMpvOsd: () => {},
  });

  const markAudio = createBuildMarkLastCardAsAudioCardMainDepsHandler({
    getAnkiIntegration: () => ({ enabled: true }),
    showMpvOsd: (text) => calls.push(`audio:${text}`),
    markLastCardAsAudioCardCore: async () => {
      calls.push('mark');
    },
  })();
  markAudio.showMpvOsd('a');
  await markAudio.markLastCardAsAudioCardCore({
    ankiIntegration: { enabled: true },
    showMpvOsd: () => {},
  });

  const mine = createBuildMineSentenceCardMainDepsHandler({
    getAnkiIntegration: () => ({ enabled: true }),
    getMpvClient: () => ({ connected: true }),
    getPrimarySubtitle: () => ({ text: '正式な字幕', startTime: 1, endTime: 3 }),
    showMpvOsd: (text) => calls.push(`mine:${text}`),
    mineSentenceCardCore: async () => true,
    recordCardsMined: (count) => calls.push(`cards:${count}`),
  })();
  assert.deepEqual(mine.getMpvClient(), { connected: true });
  assert.deepEqual(mine.getPrimarySubtitle?.(), {
    text: '正式な字幕',
    startTime: 1,
    endTime: 3,
  });
  mine.showMpvOsd('m');
  await mine.mineSentenceCardCore({
    ankiIntegration: { enabled: true },
    mpvClient: { connected: true },
    showMpvOsd: () => {},
  });
  mine.recordCardsMined(1);

  assert.deepEqual(calls, [
    'osd:x',
    'update',
    'fg:fg',
    'trigger',
    'audio:a',
    'mark',
    'mine:m',
    'cards:1',
  ]);
});
