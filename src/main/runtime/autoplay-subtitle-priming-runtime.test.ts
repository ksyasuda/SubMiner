import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createAutoplaySubtitlePrimingRuntime,
  setMpvCurrentSecondarySubText,
} from './autoplay-subtitle-priming-runtime';

test('setMpvCurrentSecondarySubText uses client setter when available', () => {
  const calls: string[] = [];
  const client = {
    currentSecondarySubText: '',
    setCurrentSecondarySubText: (text: string) => {
      calls.push(text);
    },
  };

  setMpvCurrentSecondarySubText(client, 'secondary');

  assert.deepEqual(calls, ['secondary']);
  assert.equal(client.currentSecondarySubText, '');
});

test('setMpvCurrentSecondarySubText updates client property when setter is unavailable', () => {
  const client = {
    currentSecondarySubText: '',
  };

  setMpvCurrentSecondarySubText(client, 'secondary');

  assert.equal(client.currentSecondarySubText, 'secondary');
});

test('scheduleSubtitlePrefetchRefresh logs refresh failures from timer callback', async () => {
  const logs: string[] = [];
  const runtime = createAutoplaySubtitlePrimingRuntime({
    getCurrentMediaPath: () => null,
    getMpvClient: () => null,
    setCurrentSubText: () => {},
    getCurrentSubText: () => '',
    getCurrentSubtitleData: () => null,
    setActiveParsedSubtitleMediaPath: () => {},
    subtitleProcessingController: {
      consumeCachedSubtitle: () => null,
      onSubtitleChange: () => {},
      refreshCurrentSubtitle: () => {},
    },
    emitSubtitlePayload: () => {},
    getSubtitlePrefetchService: () => null,
    getLastObservedTimePos: () => 0,
    getVisibleOverlayVisible: () => false,
    emitSecondarySubtitle: () => {},
    initSubtitlePrefetch: async () => {},
    refreshSubtitlePrefetchFromActiveTrack: async () => {
      throw new Error('refresh failed');
    },
    logDebug: (message) => logs.push(message),
  });

  runtime.scheduleSubtitlePrefetchRefresh(0);
  await new Promise((resolve) => setTimeout(resolve, 5));

  assert.deepEqual(logs, [
    '[autoplay-subtitle-prime] subtitle prefetch refresh failed: refresh failed',
  ]);
});
