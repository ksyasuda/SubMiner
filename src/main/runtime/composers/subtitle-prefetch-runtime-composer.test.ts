import assert from 'node:assert/strict';
import test from 'node:test';
import { composeSubtitlePrefetchRuntime } from './subtitle-prefetch-runtime-composer';

test('composeSubtitlePrefetchRuntime returns subtitle prefetch runtime helpers', () => {
  const composed = composeSubtitlePrefetchRuntime({
    subtitlePrefetchInitController: {
      cancelPendingInit: () => {},
      initSubtitlePrefetch: async () => {},
    },
    refreshSubtitleSidebarFromSource: async () => {},
    refreshSubtitlePrefetchFromActiveTrack: async () => {},
    scheduleSubtitlePrefetchRefresh: () => {},
    clearScheduledSubtitlePrefetchRefresh: () => {},
  });

  assert.equal(typeof composed.cancelPendingInit, 'function');
  assert.equal(typeof composed.initSubtitlePrefetch, 'function');
  assert.equal(typeof composed.refreshSubtitleSidebarFromSource, 'function');
  assert.equal(typeof composed.refreshSubtitlePrefetchFromActiveTrack, 'function');
  assert.equal(typeof composed.scheduleSubtitlePrefetchRefresh, 'function');
  assert.equal(typeof composed.clearScheduledSubtitlePrefetchRefresh, 'function');
});
