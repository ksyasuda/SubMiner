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
    logDebug: (message) => calls.push(`debug:${message}`),
  })();

  await deps.listJellyfinSubtitleTracks({} as never, {} as never, 'item');
  assert.equal(typeof deps.getMpvClient()?.requestProperty, 'function');
  deps.sendMpvCommand(['set_property', 'sid', 'auto']);
  await deps.wait(1);
  await deps.cacheSubtitleTrack({ index: 1, deliveryUrl: 'https://example.test/sub.srt' });
  deps.cleanupCachedSubtitles(['/tmp/subs']);
  deps.logDebug('oops', null);
  assert.deepEqual(calls, ['list', 'send', 'wait', 'cache', 'cleanup', 'debug:oops']);
});
