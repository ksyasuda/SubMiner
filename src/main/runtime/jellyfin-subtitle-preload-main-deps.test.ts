import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildPreloadJellyfinExternalSubtitlesMainDepsHandler } from './jellyfin-subtitle-preload-main-deps';

test('preload jellyfin external subtitles main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildPreloadJellyfinExternalSubtitlesMainDepsHandler({
    listJellyfinSubtitleTracks: async () => {
      calls.push('list');
      return [];
    },
    getMpvClient: () => ({ requestProperty: async () => [] }),
    sendMpvCommand: () => calls.push('send'),
    wait: async () => {
      calls.push('wait');
    },
    cacheSubtitleTrack: async () => {
      calls.push('cache');
      return { path: '/tmp/sub.srt', cleanupDir: '/tmp/subs' };
    },
    cleanupCachedSubtitles: () => calls.push('cleanup'),
    getSavedSubtitleDelay: (_itemId, streamIndex) => {
      calls.push(`load-delay:${streamIndex}`);
      return 1.25;
    },
    setActiveSubtitleDelayKey: (key) => calls.push(`active-delay:${key?.streamIndex ?? 'none'}`),
    loadSubtitleSourceText: async (source) => {
      calls.push(`load-source:${source}`);
      return 'subtitle';
    },
    saveSubtitleDelay: (_itemId, streamIndex, delaySeconds) => {
      calls.push(`save-delay:${streamIndex}:${delaySeconds}`);
      return true;
    },
    logDebug: (message) => calls.push(`debug:${message}`),
  })();

  await deps.listJellyfinSubtitleTracks({} as never, {} as never, 'item');
  assert.equal(typeof deps.getMpvClient()?.requestProperty, 'function');
  deps.sendMpvCommand(['set_property', 'sid', 'auto']);
  await deps.wait(1);
  await deps.cacheSubtitleTrack({ index: 1, deliveryUrl: 'https://example.test/sub.srt' });
  deps.cleanupCachedSubtitles(['/tmp/subs']);
  assert.equal(deps.getSavedSubtitleDelay?.('item', 3), 1.25);
  deps.setActiveSubtitleDelayKey?.({ itemId: 'item', streamIndex: 3 });
  assert.equal(await deps.loadSubtitleSourceText?.('/tmp/sub.srt'), 'subtitle');
  assert.equal(deps.saveSubtitleDelay?.('item', 3, -31.5), true);
  deps.logDebug('oops', null);
  assert.deepEqual(calls, [
    'list',
    'send',
    'wait',
    'cache',
    'cleanup',
    'load-delay:3',
    'active-delay:3',
    'load-source:/tmp/sub.srt',
    'save-delay:3:-31.5',
    'debug:oops',
  ]);
});
