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

test('subtitle prefetch runtime does not extract internal subtitle tracks from remote media urls', async () => {
  let extracted = false;
  const resolveSource = createResolveActiveSubtitleSidebarSourceHandler({
    getFfmpegPath: () => 'ffmpeg-custom',
    extractInternalSubtitleTrack: async () => {
      extracted = true;
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
    videoPath: 'http://jellyfin.local/Videos/movie/stream?static=true',
  });

  assert.equal(resolved, null);
  assert.equal(extracted, false);
});

test('subtitle prefetch refresh logs a warning when source resolution throws', async () => {
  const warnings: string[] = [];
  const refresh = createRefreshSubtitlePrefetchFromActiveTrackHandler({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => (name === 'path' ? '/media/video.mkv' : null),
    }),
    getLastObservedTimePos: () => 0,
    subtitlePrefetchInitController: {
      cancelPendingInit: () => {},
      initSubtitlePrefetch: async () => {},
    },
    resolveActiveSubtitleSidebarSource: async () => {
      throw new Error('ffmpeg ENOENT');
    },
    logWarn: (message) => warnings.push(message),
  });

  await refresh();

  assert.equal(warnings.length, 1);
  assert.match(warnings[0]!, /\[subtitle-prefetch\].*ffmpeg ENOENT/);
});

test('subtitle prefetch refresh logs debug when mpv client is not connected', async () => {
  const debugs: string[] = [];
  const refresh = createRefreshSubtitlePrefetchFromActiveTrackHandler({
    getMpvClient: () => null,
    getLastObservedTimePos: () => 0,
    subtitlePrefetchInitController: {
      cancelPendingInit: () => {},
      initSubtitlePrefetch: async () => {},
    },
    resolveActiveSubtitleSidebarSource: async () => null,
    logDebug: (message) => debugs.push(message),
  });

  await refresh();

  assert.equal(debugs.length, 1);
  assert.match(debugs[0]!, /\[subtitle-prefetch\].*not connected/);
});

test('subtitle prefetch refresh logs debug when no subtitle source resolves', async () => {
  const debugs: string[] = [];
  const cancels: number[] = [];
  const refresh = createRefreshSubtitlePrefetchFromActiveTrackHandler({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => (name === 'path' ? '/media/video.mkv' : null),
    }),
    getLastObservedTimePos: () => 0,
    subtitlePrefetchInitController: {
      cancelPendingInit: () => {
        cancels.push(1);
      },
      initSubtitlePrefetch: async () => {},
    },
    resolveActiveSubtitleSidebarSource: async () => null,
    logDebug: (message) => debugs.push(message),
  });

  await refresh();

  assert.deepEqual(cancels, [1]);
  assert.equal(debugs.length, 1);
  assert.match(debugs[0]!, /\[subtitle-prefetch\].*no active subtitle source/);
});

test('subtitle source resolver logs debug when internal track extraction is unavailable', async () => {
  const debugs: string[] = [];
  const resolveSource = createResolveActiveSubtitleSidebarSourceHandler({
    getFfmpegPath: () => 'ffmpeg',
    extractInternalSubtitleTrack: async () => null,
    logDebug: (message) => debugs.push(message),
  });

  const resolved = await resolveSource({
    currentExternalFilenameRaw: null,
    currentTrackRaw: {
      type: 'sub',
      id: 3,
      'ff-index': 7,
      codec: 'hdmv_pgs_subtitle',
    },
    trackListRaw: [],
    sidRaw: 3,
    videoPath: '/media/video.mkv',
  });

  assert.equal(resolved, null);
  assert.equal(debugs.length, 1);
  assert.match(debugs[0]!, /\[subtitle-prefetch\].*extraction.*hdmv_pgs_subtitle/);
});

test('subtitle source resolver logs debug when no active subtitle track is selected', async () => {
  const debugs: string[] = [];
  const resolveSource = createResolveActiveSubtitleSidebarSourceHandler({
    getFfmpegPath: () => 'ffmpeg',
    extractInternalSubtitleTrack: async () => {
      throw new Error('should not extract without a track');
    },
    logDebug: (message) => debugs.push(message),
  });

  const resolved = await resolveSource({
    currentExternalFilenameRaw: null,
    currentTrackRaw: null,
    trackListRaw: [],
    sidRaw: null,
    videoPath: '/media/video.mkv',
  });

  assert.equal(resolved, null);
  assert.equal(debugs.length, 1);
  assert.match(debugs[0]!, /\[subtitle-prefetch\].*no active subtitle track/);
});

test('subtitle source resolver does not fall back to the primary selected track for secondary', async () => {
  const resolveSource = createResolveActiveSubtitleSidebarSourceHandler({
    getFfmpegPath: () => 'ffmpeg',
    extractInternalSubtitleTrack: async () => {
      throw new Error('should not extract the primary track');
    },
  });

  const resolved = await resolveSource({
    currentExternalFilenameRaw: null,
    currentTrackRaw: null,
    trackListRaw: [
      {
        type: 'sub',
        id: 1,
        selected: true,
        external: true,
        'external-filename': '/subs/primary.ass',
      },
    ],
    sidRaw: null,
    videoPath: '/media/video.mkv',
    allowSelectedFallback: false,
  });

  assert.equal(resolved, null);
});
