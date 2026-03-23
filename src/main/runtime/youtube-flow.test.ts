import assert from 'node:assert/strict';
import test from 'node:test';
import { createYoutubeFlowRuntime } from './youtube-flow';
import type { YoutubePickerOpenPayload, YoutubeTrackOption } from '../../types';

const primaryTrack: YoutubeTrackOption = {
  id: 'auto:ja-orig',
  language: 'ja',
  sourceLanguage: 'ja-orig',
  kind: 'auto',
  label: 'Japanese (auto)',
};

const secondaryTrack: YoutubeTrackOption = {
  id: 'manual:en',
  language: 'en',
  sourceLanguage: 'en',
  kind: 'manual',
  label: 'English (manual)',
};

test('youtube flow clears internal tracks and binds external primary+secondary subtitles', async () => {
  const commands: Array<Array<string | number>> = [];
  const osdMessages: string[] = [];
  const order: string[] = [];
  const refreshedSubtitles: string[] = [];
  const waits: number[] = [];
  const focusOverlayCalls: string[] = [];
  let pickerPayload: YoutubePickerOpenPayload | null = null;
  let trackListRequests = 0;

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack, secondaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async ({ tracks }) => {
      assert.deepEqual(
        tracks.map((track) => track.id),
        [primaryTrack.id, secondaryTrack.id],
      );
      return new Map<string, string>([
        [primaryTrack.id, '/tmp/auto-ja-orig.vtt'],
        [secondaryTrack.id, '/tmp/manual-en.vtt'],
      ]);
    },
    acquireYoutubeSubtitleTrack: async ({ track }) => {
      if (track.id === primaryTrack.id) {
        return { path: '/tmp/auto-ja-orig.vtt' };
      }
      return { path: '/tmp/manual-en.vtt' };
    },
    retimeYoutubePrimaryTrack: async ({ primaryPath, secondaryPath }) => {
      assert.equal(primaryPath, '/tmp/auto-ja-orig.vtt');
      assert.equal(secondaryPath, '/tmp/manual-en.vtt');
      return '/tmp/auto-ja-orig_retimed.vtt';
    },
    startTokenizationWarmups: async () => {
      order.push('start-tokenization-warmups');
    },
    waitForTokenizationReady: async () => {
      order.push('wait-tokenization-ready');
    },
    waitForAnkiReady: async () => {
      order.push('wait-anki-ready');
    },
    waitForPlaybackWindowReady: async () => {
      order.push('wait-window-ready');
    },
    waitForOverlayGeometryReady: async () => {
      order.push('wait-overlay-geometry');
    },
    focusOverlayWindow: () => {
      focusOverlayCalls.push('focus-overlay');
    },
    openPicker: async (payload) => {
      assert.deepEqual(waits, [150]);
      order.push('open-picker');
      pickerPayload = payload;
      queueMicrotask(() => {
        void runtime.resolveActivePicker({
          sessionId: payload.sessionId,
          action: 'use-selected',
          primaryTrackId: primaryTrack.id,
          secondaryTrackId: secondaryTrack.id,
        });
      });
      return true;
    },
    pauseMpv: () => {
      commands.push(['set_property', 'pause', 'yes']);
    },
    resumeMpv: () => {
      commands.push(['set_property', 'pause', 'no']);
    },
    sendMpvCommand: (command) => {
      commands.push(command);
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      assert.equal(name, 'track-list');
      trackListRequests += 1;
      if (trackListRequests === 1) {
        return [{ type: 'sub', id: 1, lang: 'ja', external: false, title: 'internal' }];
      }
      return [
        {
          type: 'sub',
          id: 5,
          lang: 'ja-orig',
          title: 'primary',
          external: true,
          'external-filename': '/tmp/auto-ja-orig_retimed.vtt',
        },
        {
          type: 'sub',
          id: 6,
          lang: 'en',
          title: 'secondary',
          external: true,
          'external-filename': '/tmp/manual-en.vtt',
        },
      ];
    },
    refreshCurrentSubtitle: (text) => {
      refreshedSubtitles.push(text);
    },
    wait: async (ms) => {
      waits.push(ms);
    },
    showMpvOsd: (text) => {
      osdMessages.push(text);
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });
  await runtime.runYoutubePlaybackFlow({ url: 'https://example.com', mode: 'download' });

  assert.ok(pickerPayload);
  assert.deepEqual(order, [
    'start-tokenization-warmups',
    'wait-window-ready',
    'wait-overlay-geometry',
    'open-picker',
    'wait-tokenization-ready',
    'wait-anki-ready',
  ]);
  assert.deepEqual(osdMessages, [
    'Opening YouTube video',
    'Getting subtitles...',
    'Downloading subtitles...',
    'Loading subtitles...',
    'Primary and secondary subtitles loaded.',
  ]);
  assert.deepEqual(commands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'sub-delay', 0],
    ['set_property', 'sid', 'no'],
    ['set_property', 'secondary-sid', 'no'],
    ['sub-add', '/tmp/auto-ja-orig_retimed.vtt', 'select', 'auto-ja-orig_retimed.vtt', 'ja-orig'],
    ['sub-add', '/tmp/manual-en.vtt', 'cached', 'manual-en.vtt', 'en'],
    ['set_property', 'sid', 5],
    ['set_property', 'secondary-sid', 6],
    ['script-message', 'subminer-autoplay-ready'],
    ['set_property', 'pause', 'no'],
  ]);
  assert.deepEqual(focusOverlayCalls, ['focus-overlay']);
  assert.deepEqual(refreshedSubtitles, ['字幕です']);
});

test('youtube flow can cancel active picker session', async () => {
  const focusOverlayCalls: string[] = [];
  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () => {
      throw new Error('should not batch download after cancel');
    },
    acquireYoutubeSubtitleTrack: async () => {
      throw new Error('should not download after cancel');
    },
    retimeYoutubePrimaryTrack: async ({ primaryPath }) => primaryPath,
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {
      focusOverlayCalls.push('focus-overlay');
    },
    openPicker: async (payload) => {
      queueMicrotask(() => {
        assert.equal(runtime.cancelActivePicker(), true);
        void runtime.resolveActivePicker({
          sessionId: payload.sessionId,
          action: 'use-selected',
          primaryTrackId: primaryTrack.id,
          secondaryTrackId: null,
        });
      });
      return true;
    },
    pauseMpv: () => {},
    resumeMpv: () => {},
    sendMpvCommand: () => {},
    requestMpvProperty: async () => null,
    refreshCurrentSubtitle: () => {},
    wait: async () => {},
    showMpvOsd: () => {},
    warn: () => {},
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.runYoutubePlaybackFlow({ url: 'https://example.com', mode: 'download' });
  assert.equal(runtime.hasActiveSession(), false);
  assert.equal(runtime.cancelActivePicker(), false);
  assert.deepEqual(focusOverlayCalls, ['focus-overlay']);
});

test('youtube flow retries secondary after partial batch subtitle failure', async () => {
  const acquireSingleCalls: string[] = [];
  const commands: Array<Array<string | number>> = [];
  const focusOverlayCalls: string[] = [];
  const refreshedSubtitles: string[] = [];
  const warns: string[] = [];
  const waits: number[] = [];
  let trackListRequests = 0;

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack, secondaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () => {
      return new Map<string, string>([[primaryTrack.id, '/tmp/auto-ja-orig.vtt']]);
    },
    acquireYoutubeSubtitleTrack: async ({ track }) => {
      acquireSingleCalls.push(track.id);
      return { path: `/tmp/${track.id}.vtt` };
    },
    retimeYoutubePrimaryTrack: async ({ primaryPath, secondaryPath }) => {
      assert.equal(primaryPath, '/tmp/auto-ja-orig.vtt');
      assert.equal(secondaryPath, '/tmp/manual:en.vtt');
      return primaryPath;
    },
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {
      focusOverlayCalls.push('focus-overlay');
    },
    openPicker: async (payload) => {
      queueMicrotask(() => {
        void runtime.resolveActivePicker({
          sessionId: payload.sessionId,
          action: 'use-selected',
          primaryTrackId: primaryTrack.id,
          secondaryTrackId: secondaryTrack.id,
        });
      });
      return true;
    },
    pauseMpv: () => {},
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      commands.push(command);
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      assert.equal(name, 'track-list');
      trackListRequests += 1;
      if (trackListRequests === 1) {
        return [
          {
            type: 'sub',
            id: 5,
            lang: 'ja-orig',
            title: 'primary',
            external: true,
            'external-filename': '/tmp/auto-ja-orig.vtt',
          },
        ];
      }
      return [
        {
          type: 'sub',
          id: 5,
          lang: 'ja-orig',
          title: 'primary',
          external: true,
          'external-filename': '/tmp/auto-ja-orig.vtt',
        },
        {
          type: 'sub',
          id: 6,
          lang: 'en',
          title: 'secondary',
          external: true,
          'external-filename': '/tmp/manual:en.vtt',
        },
      ];
    },
    refreshCurrentSubtitle: (text) => {
      refreshedSubtitles.push(text);
    },
    wait: async (ms) => {
      waits.push(ms);
    },
    showMpvOsd: () => {},
    warn: (message) => {
      warns.push(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.runYoutubePlaybackFlow({ url: 'https://example.com', mode: 'download' });

  assert.deepEqual(acquireSingleCalls, [secondaryTrack.id]);
  assert.ok(waits.includes(150));
  assert.deepEqual(focusOverlayCalls, ['focus-overlay']);
  assert.deepEqual(refreshedSubtitles, ['字幕です']);
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'sub-add' &&
        command[1] === '/tmp/manual:en.vtt' &&
        command[2] === 'cached',
    ),
  );
  assert.equal(warns.length, 0);
});

test('youtube flow waits for tokenization readiness before releasing playback', async () => {
  const commands: Array<Array<string | number>> = [];
  const releaseOrder: string[] = [];
  let tokenizationReadyRegistered = false;
  let resolveTokenizationReady: () => void = () => {
    throw new Error('expected tokenization readiness waiter');
  };

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () => new Map<string, string>(),
    acquireYoutubeSubtitleTrack: async () => ({ path: '/tmp/auto-ja-orig.vtt' }),
    retimeYoutubePrimaryTrack: async ({ primaryPath }) => primaryPath,
    startTokenizationWarmups: async () => {
      releaseOrder.push('start-warmups');
    },
    waitForTokenizationReady: async () => {
      releaseOrder.push('wait-tokenization-ready:start');
      await new Promise<void>((resolve) => {
        tokenizationReadyRegistered = true;
        resolveTokenizationReady = resolve;
      });
      releaseOrder.push('wait-tokenization-ready:end');
    },
    waitForAnkiReady: async () => {
      releaseOrder.push('wait-anki-ready');
    },
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {
      releaseOrder.push('focus-overlay');
    },
    openPicker: async (payload) => {
      queueMicrotask(() => {
        void runtime.resolveActivePicker({
          sessionId: payload.sessionId,
          action: 'use-selected',
          primaryTrackId: primaryTrack.id,
          secondaryTrackId: null,
        });
      });
      return true;
    },
    pauseMpv: () => {},
    resumeMpv: () => {
      commands.push(['set_property', 'pause', 'no']);
      releaseOrder.push('resume');
    },
    sendMpvCommand: (command) => {
      commands.push(command);
      if (command[0] === 'script-message' && command[1] === 'subminer-autoplay-ready') {
        releaseOrder.push('autoplay-ready');
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      return [
        {
          type: 'sub',
          id: 5,
          lang: 'ja-orig',
          title: 'primary',
          external: true,
          'external-filename': '/tmp/auto-ja-orig.vtt',
        },
      ];
    },
    refreshCurrentSubtitle: () => {},
    wait: async () => {},
    showMpvOsd: () => {},
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  const flowPromise = runtime.runYoutubePlaybackFlow({ url: 'https://example.com', mode: 'download' });
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(tokenizationReadyRegistered, true);
  assert.deepEqual(releaseOrder, ['start-warmups', 'wait-tokenization-ready:start']);
  assert.equal(commands.some((command) => command[1] === 'subminer-autoplay-ready'), false);

  resolveTokenizationReady();
  await flowPromise;

  assert.deepEqual(releaseOrder, [
    'start-warmups',
    'wait-tokenization-ready:start',
    'wait-tokenization-ready:end',
    'wait-anki-ready',
    'autoplay-ready',
    'resume',
    'focus-overlay',
  ]);
});
