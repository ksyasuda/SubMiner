import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { createSubtitleRuntime } from './subtitle-runtime';

function createResolvedConfig() {
  return {
    subtitleStyle: {
      frequencyDictionary: {
        enabled: true,
        topX: 5,
        mode: 'top',
      },
    },
    subtitleSidebar: {
      autoScroll: true,
      pauseVideoOnHover: false,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: '#111111',
      textColor: '#ffffff',
      fontFamily: 'sans-serif',
      fontSize: 24,
      timestampColor: '#cccccc',
      activeLineColor: '#ffffff',
      activeLineBackgroundColor: '#222222',
      hoverLineBackgroundColor: '#333333',
    },
    subtitlePosition: {
      yPercent: 84,
    },
    subsync: {
      defaultMode: 'auto' as const,
      alass_path: 'alass',
      ffmpeg_path: 'ffmpeg',
      ffsubsync_path: 'ffsubsync',
      replace: false,
    },
  } as never;
}

function createMpvClient(properties: Record<string, unknown>) {
  return {
    connected: true,
    currentSubStart: 1.25,
    currentSubEnd: 2.5,
    currentTimePos: 12.5,
    requestProperty: async (name: string) => properties[name],
  };
}

function createRuntime(overrides: Partial<Parameters<typeof createSubtitleRuntime>[0]> = {}) {
  const calls: string[] = [];
  const config = createResolvedConfig();
  let subtitlePosition: unknown = null;
  let pendingSubtitlePosition: unknown = null;
  const runtime = createSubtitleRuntime({
    getResolvedConfig: () => config,
    getCurrentMediaPath: () => '/media/episode.mkv',
    getCurrentMediaTitle: () => 'Episode',
    getCurrentSubText: () => 'current subtitle',
    getCurrentSubAssText: () => '[Events]',
    getMpvClient: () =>
      createMpvClient({
        'current-tracks/sub/external-filename': '/tmp/episode.ass',
        'current-tracks/sub': {
          type: 'sub',
          id: 3,
          external: true,
          'external-filename': '/tmp/episode.ass',
        },
        'track-list': [
          {
            type: 'sub',
            id: 3,
            external: true,
            'external-filename': '/tmp/episode.ass',
          },
        ],
        sid: 3,
        path: '/media/episode.mkv',
      }),
    broadcastToOverlayWindows: (channel, payload) => {
      calls.push(`${channel}:${JSON.stringify(payload)}`);
    },
    subtitleWsService: {
      broadcast: () => calls.push('subtitle-ws'),
    },
    annotationSubtitleWsService: {
      broadcast: () => calls.push('annotation-ws'),
    },
    subtitlePositionsDir: fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subtitle-runtime-')),
    setSubtitlePosition: (position) => {
      subtitlePosition = position;
    },
    setPendingSubtitlePosition: (position) => {
      pendingSubtitlePosition = position;
    },
    clearPendingSubtitlePosition: () => {
      pendingSubtitlePosition = null;
    },
    parseSubtitleCues: (content) => [
      {
        startTime: 0,
        endTime: 1,
        text: content.trim(),
      },
    ],
    createSubtitlePrefetchService: ({ cues }) => ({
      start: () => calls.push(`start:${cues.length}`),
      stop: () => calls.push('stop'),
      onSeek: (time) => calls.push(`seek:${time}`),
      pause: () => calls.push('pause'),
      resume: () => calls.push('resume'),
    }),
    logDebug: (message) => calls.push(`debug:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    schedule: (fn, delayMs) => setTimeout(fn, delayMs),
    clearSchedule: (timer) => clearTimeout(timer),
    ...overrides,
  });
  return { runtime, calls, subtitlePosition, pendingSubtitlePosition };
}

test('subtitle runtime schedules and cancels subtitle prefetch refreshes', async () => {
  const calls: string[] = [];
  const { runtime } = createRuntime({
    refreshSubtitlePrefetchFromActiveTrack: async () => {
      calls.push('refresh');
    },
  });

  runtime.scheduleSubtitlePrefetchRefresh(5);
  runtime.clearScheduledSubtitlePrefetchRefresh();
  await new Promise((resolve) => setTimeout(resolve, 20));

  assert.deepEqual(calls, []);
});

test('subtitle runtime times out remote subtitle source fetches', async () => {
  const { runtime } = createRuntime({
    fetchImpl: async (_url, init) => {
      await new Promise<void>((_resolve, reject) => {
        init?.signal?.addEventListener('abort', () => reject(new Error('aborted')), { once: true });
      });
      return new Response('');
    },
    subtitleSourceFetchTimeoutMs: 10,
  });

  await assert.rejects(
    async () => await runtime.loadSubtitleSourceText('https://example.com/subtitles.srt'),
    /aborted/,
  );
});

test('subtitle runtime reuses cached sidebar cues for the same source key', async () => {
  const subtitlePath = path.join(
    fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subtitle-cache-')),
    'episode.ass',
  );
  fs.writeFileSync(
    subtitlePath,
    `1
00:00:01,000 --> 00:00:02,000
Hello`,
  );

  let loadCount = 0;
  const { runtime } = createRuntime({
    getMpvClient: () =>
      createMpvClient({
        'current-tracks/sub/external-filename': subtitlePath,
        'current-tracks/sub': {
          type: 'sub',
          id: 3,
          external: true,
          'external-filename': subtitlePath,
        },
        'track-list': [
          {
            type: 'sub',
            id: 3,
            external: true,
            'external-filename': subtitlePath,
          },
        ],
        sid: 3,
        path: '/media/episode.mkv',
      }),
    loadSubtitleSourceText: async () => {
      loadCount += 1;
      return fs.readFileSync(subtitlePath, 'utf8');
    },
  });

  const first = await runtime.getSubtitleSidebarSnapshot();
  const second = await runtime.getSubtitleSidebarSnapshot();

  assert.equal(loadCount, 1);
  assert.deepEqual(second.cues, first.cues);
  assert.equal(second.currentSubtitle.text, 'current subtitle');
});
