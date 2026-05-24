import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createRefreshSubtitlePrefetchFromActiveTrackHandler,
  createResolveActiveSubtitleSidebarSourceHandler,
} from './subtitle-prefetch-runtime';

test('subtitle prefetch runtime resolves direct external subtitle sources first', async () => {
  const resolveSource = createResolveActiveSubtitleSidebarSourceHandler({
    getFfmpegPath: () => 'ffmpeg',
    extractInternalSubtitleTrack: async () => {
      throw new Error('should not extract external tracks');
    },
  });

  const resolved = await resolveSource({
    currentExternalFilenameRaw: ' /tmp/current.ass ',
    currentTrackRaw: null,
    trackListRaw: null,
    sidRaw: null,
    videoPath: '/media/video.mkv',
  });

  assert.deepEqual(resolved, {
    path: '/tmp/current.ass',
    sourceKey: '/tmp/current.ass',
  });
});

test('subtitle prefetch runtime extracts internal subtitle tracks into a stable source key', async () => {
  const resolveSource = createResolveActiveSubtitleSidebarSourceHandler({
    getFfmpegPath: () => 'ffmpeg-custom',
    extractInternalSubtitleTrack: async (ffmpegPath, videoPath, track) => {
      assert.equal(ffmpegPath, 'ffmpeg-custom');
      assert.equal(videoPath, '/media/video.mkv');
      assert.equal((track as Record<string, unknown>)['ff-index'], 7);
      return {
        path: '/tmp/subminer-sidebar-123/track_7.ass',
        cleanup: async () => {},
      };
    },
  });

  const resolved = await resolveSource({
    currentExternalFilenameRaw: null,
    currentTrackRaw: {
      type: 'sub',
      id: 3,
      'ff-index': 7,
      codec: 'ass',
    },
    trackListRaw: [],
    sidRaw: 3,
    videoPath: '/media/video.mkv',
  });

  assert.deepEqual(resolved, {
    path: '/tmp/subminer-sidebar-123/track_7.ass',
    sourceKey: 'internal:/media/video.mkv:track:3:ff:7',
    cleanup: resolved?.cleanup,
  });
});

test('subtitle prefetch runtime preserves parsed cues when YouTube active track source is unresolved', async () => {
  const calls: string[] = [];
  const refresh = createRefreshSubtitlePrefetchFromActiveTrackHandler({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'path') return 'https://www.youtube.com/watch?v=video123';
        if (name === 'track-list') {
          return [
            {
              type: 'sub',
              id: 4,
              lang: 'ja',
              title: 'Japanese',
              external: true,
            },
          ];
        }
        if (name === 'sid') return 4;
        return null;
      },
    }),
    getLastObservedTimePos: () => 12,
    subtitlePrefetchInitController: {
      cancelPendingInit: () => {
        calls.push('cancel');
      },
      initSubtitlePrefetch: async () => {
        calls.push('init');
      },
    },
    resolveActiveSubtitleSidebarSource: async () => null,
    shouldKeepExistingCuesOnMissingSource: (videoPath) => videoPath.includes('youtube.com'),
  });

  await refresh();

  assert.deepEqual(calls, []);
});
