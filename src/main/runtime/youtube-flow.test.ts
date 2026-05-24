import assert from 'node:assert/strict';
import path from 'node:path';
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

test('youtube flow can open a manual picker session and load the selected subtitles', async () => {
  const commands: Array<Array<string | number>> = [];
  const focusOverlayCalls: string[] = [];
  const osdMessages: string[] = [];
  const openedPayloads: YoutubePickerOpenPayload[] = [];
  const waits: number[] = [];
  const refreshedSidebarSources: string[] = [];

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
    acquireYoutubeSubtitleTrack: async ({ track }) => ({
      path: `/tmp/${track.id.replace(/[^a-z0-9_-]+/gi, '-')}.vtt`,
    }),
    openPicker: async (payload) => {
      openedPayloads.push(payload);
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
          'external-filename': '/tmp/manual-en.vtt',
        },
      ];
    },
    refreshCurrentSubtitle: () => {},
    refreshSubtitleSidebarSource: async (sourcePath: string) => {
      refreshedSidebarSources.push(sourcePath);
    },
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async (ms) => {
      waits.push(ms);
    },
    waitForPlaybackWindowReady: async () => {
      waits.push(1);
    },
    waitForOverlayGeometryReady: async () => {
      waits.push(2);
    },
    focusOverlayWindow: () => {
      focusOverlayCalls.push('focus-overlay');
    },
    showMpvOsd: (text) => {
      osdMessages.push(text);
    },
    reportSubtitleFailure: () => {
      throw new Error('manual picker success should not report failure');
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.equal(openedPayloads.length, 1);
  assert.equal(openedPayloads[0]?.defaultPrimaryTrackId, primaryTrack.id);
  assert.equal(openedPayloads[0]?.defaultSecondaryTrackId, secondaryTrack.id);
  assert.ok(waits.includes(150));
  assert.deepEqual(osdMessages, [
    'Getting subtitles...',
    'Downloading subtitles...',
    'Loading subtitles...',
    'Primary and secondary subtitles loaded.',
  ]);
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'sub-add' &&
        command[1] === '/tmp/auto-ja-orig.vtt' &&
        command[2] === 'select',
    ),
  );
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'set_property' && command[1] === 'sub-visibility' && command[2] === 'yes',
    ),
  );
  assert.ok(
    commands.every(
      (command) =>
        !(
          command[0] === 'set_property' &&
          command[1] === 'secondary-sub-visibility' &&
          command[2] === 'yes'
        ),
    ),
  );
  assert.deepEqual(refreshedSidebarSources, ['/tmp/auto-ja-orig.vtt']);
  assert.deepEqual(focusOverlayCalls, ['focus-overlay']);
});

test('youtube flow retries secondary after partial batch subtitle failure', async () => {
  const acquireSingleCalls: string[] = [];
  const commands: Array<Array<string | number>> = [];
  const waits: number[] = [];
  let secondaryTrackAdded = false;

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack, secondaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () =>
      new Map<string, string>([[primaryTrack.id, '/tmp/auto-ja-orig.vtt']]),
    acquireYoutubeSubtitleTrack: async ({ track }) => {
      acquireSingleCalls.push(track.id);
      return { path: `/tmp/${track.id}.vtt` };
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
      if (
        command[0] === 'sub-add' &&
        command[1] === '/tmp/manual:en.vtt' &&
        command[2] === 'cached'
      ) {
        secondaryTrackAdded = true;
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      return secondaryTrackAdded
        ? [
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
          ]
        : [
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
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async (ms) => {
      waits.push(ms);
    },
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: () => {
      throw new Error('secondary retry should not report primary failure');
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.deepEqual(acquireSingleCalls, [secondaryTrack.id]);
  assert.ok(waits.includes(350));
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'sub-add' && command[1] === '/tmp/manual:en.vtt' && command[2] === 'cached',
    ),
  );
});

test('youtube flow reports probe failure through the configured reporter in manual mode', async () => {
  const failures: string[] = [];

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => {
      throw new Error('probe failed');
    },
    acquireYoutubeSubtitleTracks: async () => new Map(),
    acquireYoutubeSubtitleTrack: async () => ({ path: '/tmp/unused.vtt' }),
    openPicker: async () => true,
    pauseMpv: () => {},
    resumeMpv: () => {},
    sendMpvCommand: () => {},
    requestMpvProperty: async () => null,
    refreshCurrentSubtitle: () => {},
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: (message) => {
      failures.push(message);
    },
    warn: () => {},
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.deepEqual(failures, [
    'Primary subtitles failed to load. Use the YouTube subtitle picker to try manually.',
  ]);
});

test('youtube flow does not report failure when subtitle track binds before cue text appears', async () => {
  const failures: string[] = [];
  const loadedSignals: string[] = [];

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () => new Map(),
    acquireYoutubeSubtitleTrack: async () => ({ path: '/tmp/auto-ja-orig.vtt' }),
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
    resumeMpv: () => {},
    sendMpvCommand: () => {},
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '';
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
    refreshCurrentSubtitle: () => {
      throw new Error('should not refresh empty subtitle text');
    },
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: (message) => {
      failures.push(message);
    },
    notifyPrimarySubtitleLoaded: () => {
      loadedSignals.push('loaded');
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.deepEqual(failures, []);
  assert.deepEqual(loadedSignals, ['loaded']);
});

test('youtube flow does not fail when mpv reports sub-text as unavailable after track bind', async () => {
  const failures: string[] = [];

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () => new Map(),
    acquireYoutubeSubtitleTrack: async () => ({ path: '/tmp/auto-ja-orig.vtt' }),
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
    resumeMpv: () => {},
    sendMpvCommand: () => {},
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        throw new Error("Failed to read MPV property 'sub-text': property unavailable");
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
    refreshCurrentSubtitle: () => {
      throw new Error('should not refresh when sub-text is unavailable');
    },
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: (message) => {
      failures.push(message);
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.deepEqual(failures, []);
});

test('youtube flow retries secondary subtitle selection until mpv reports the expected secondary sid', async () => {
  const commands: Array<Array<string | number>> = [];
  const waits: number[] = [];
  let secondarySidReads = 0;

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack, secondaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () =>
      new Map<string, string>([
        [primaryTrack.id, '/tmp/auto-ja-orig.vtt'],
        [secondaryTrack.id, '/tmp/manual-en.vtt'],
      ]),
    acquireYoutubeSubtitleTrack: async ({ track }) => ({
      path: `/tmp/${track.id.replace(/[^a-z0-9_-]+/gi, '-')}.vtt`,
    }),
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
      if (name === 'secondary-sid') {
        secondarySidReads += 1;
        return secondarySidReads >= 2 ? 6 : null;
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
          title: 'manual-en.vtt',
          external: true,
          'external-filename': null,
        },
      ];
    },
    refreshCurrentSubtitle: () => {},
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async (ms) => {
      waits.push(ms);
    },
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: () => {
      throw new Error('secondary selection retry should not report failure');
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.equal(
    commands.filter(
      (command) =>
        command[0] === 'set_property' && command[1] === 'secondary-sid' && command[2] === 6,
    ).length,
    2,
  );
  assert.ok(waits.includes(100));
});

test('youtube flow reuses the matching existing manual secondary track instead of a loose language match', async () => {
  const commands: Array<Array<string | number>> = [];
  let selectedSecondarySid: number | null = null;

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [
        primaryTrack,
        {
          ...secondaryTrack,
          id: 'manual:en',
          sourceLanguage: 'en',
          kind: 'manual',
          title: 'manual-en.vtt',
        },
      ],
    }),
    acquireYoutubeSubtitleTracks: async () =>
      new Map<string, string>([
        [primaryTrack.id, '/tmp/auto-ja-orig.vtt'],
        [secondaryTrack.id, '/tmp/manual-en.vtt'],
      ]),
    acquireYoutubeSubtitleTrack: async ({ track }) => ({
      path: `/tmp/${track.id.replace(/[^a-z0-9_-]+/gi, '-')}.vtt`,
    }),
    openPicker: async (payload) => {
      queueMicrotask(() => {
        void runtime.resolveActivePicker({
          sessionId: payload.sessionId,
          action: 'use-selected',
          primaryTrackId: primaryTrack.id,
          secondaryTrackId: 'manual:en',
        });
      });
      return true;
    },
    pauseMpv: () => {},
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      commands.push(command);
      if (
        command[0] === 'set_property' &&
        command[1] === 'secondary-sid' &&
        typeof command[2] === 'number'
      ) {
        selectedSecondarySid = command[2];
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      if (name === 'sid') {
        return 5;
      }
      if (name === 'secondary-sid') {
        return selectedSecondarySid;
      }
      return [
        {
          type: 'sub',
          id: 5,
          lang: 'ja-orig',
          title: 'auto-ja-orig.vtt',
          external: true,
          'external-filename': '/tmp/auto-ja-orig.vtt',
        },
        {
          type: 'sub',
          id: 6,
          lang: 'en',
          title: 'English',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 8,
          lang: 'en',
          title: 'manual-en.vtt',
          external: true,
          'external-filename': null,
        },
      ];
    },
    refreshCurrentSubtitle: () => {},
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: () => {
      throw new Error('authoritative secondary bind should not report failure');
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.equal(selectedSecondarySid, 8);
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'set_property' && command[1] === 'secondary-sid' && command[2] === 8,
    ),
  );
});

test('youtube flow leaves non-authoritative youtube subtitle tracks untouched after authoritative tracks bind', async () => {
  const commands: Array<Array<string | number>> = [];
  let selectedPrimarySid: number | null = null;
  let selectedSecondarySid: number | null = null;

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack, secondaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () =>
      new Map<string, string>([
        [primaryTrack.id, '/tmp/manual-ja.ja.srt'],
        [secondaryTrack.id, '/tmp/manual-en.en.srt'],
      ]),
    acquireYoutubeSubtitleTrack: async ({ track }) => ({
      path: `/tmp/${track.id.replace(/[^a-z0-9_-]+/gi, '-')}.vtt`,
    }),
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
      if (command[0] === 'set_property' && command[1] === 'sid' && typeof command[2] === 'number') {
        selectedPrimarySid = command[2];
      }
      if (
        command[0] === 'set_property' &&
        command[1] === 'secondary-sid' &&
        typeof command[2] === 'number'
      ) {
        selectedSecondarySid = command[2];
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      if (name === 'sid') {
        return selectedPrimarySid;
      }
      if (name === 'secondary-sid') {
        return selectedSecondarySid;
      }
      return [
        {
          type: 'sub',
          id: 1,
          lang: 'en',
          title: 'English',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 2,
          lang: 'ja',
          title: 'Japanese',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 3,
          lang: 'ja-en',
          title: 'Japanese from English',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 4,
          lang: 'ja-ja',
          title: 'Japanese from Japanese',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 5,
          lang: 'ja-orig',
          title: 'auto-ja-orig.vtt',
          external: true,
          'external-filename': '/tmp/auto-ja-orig.vtt',
        },
        {
          type: 'sub',
          id: 6,
          lang: 'en',
          title: 'manual-en.en.srt',
          external: true,
          'external-filename': '/tmp/manual-en.en.srt',
        },
      ];
    },
    refreshCurrentSubtitle: () => {},
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: () => {
      throw new Error('authoritative bind should not report failure');
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.equal(
    commands.some((command) => command[0] === 'sub-remove'),
    false,
  );
});

test('youtube flow injects downloaded primary while reusing existing manual secondary tracks', async () => {
  const commands: Array<Array<string | number>> = [];
  let selectedPrimarySid: number | null = null;
  let selectedSecondarySid: number | null = null;
  let downloadedPrimaryAdded = false;
  const refreshedSidebarSources: string[] = [];
  const downloadedPrimaryPath = '/tmp/manual-ja.ja.srt';

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [
        {
          ...primaryTrack,
          id: 'manual:ja',
          sourceLanguage: 'ja',
          kind: 'manual',
          title: 'Japanese',
        },
        {
          ...secondaryTrack,
          id: 'manual:en',
          sourceLanguage: 'en',
          kind: 'manual',
          title: 'English',
        },
      ],
    }),
    acquireYoutubeSubtitleTracks: async () => {
      throw new Error('should not batch download when both manual tracks already exist in mpv');
    },
    acquireYoutubeSubtitleTrack: async ({ track }) => {
      if (track.language === 'ja') {
        return { path: downloadedPrimaryPath };
      }
      throw new Error('should not download secondary track when manual english already exists');
    },
    openPicker: async (payload) => {
      queueMicrotask(() => {
        void runtime.resolveActivePicker({
          sessionId: payload.sessionId,
          action: 'use-selected',
          primaryTrackId: 'manual:ja',
          secondaryTrackId: 'manual:en',
        });
      });
      return true;
    },
    pauseMpv: () => {},
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      commands.push(command);
      if (
        command[0] === 'sub-add' &&
        command[1] === downloadedPrimaryPath &&
        command[2] === 'select'
      ) {
        downloadedPrimaryAdded = true;
      }
      if (command[0] === 'set_property' && command[1] === 'sid' && typeof command[2] === 'number') {
        selectedPrimarySid = command[2];
      }
      if (
        command[0] === 'set_property' &&
        command[1] === 'secondary-sid' &&
        typeof command[2] === 'number'
      ) {
        selectedSecondarySid = command[2];
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      if (name === 'sid') {
        return selectedPrimarySid;
      }
      if (name === 'secondary-sid') {
        return selectedSecondarySid;
      }
      const tracks: Array<Record<string, unknown>> = [
        {
          type: 'sub',
          id: 1,
          lang: 'en',
          title: 'English',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 2,
          lang: 'ja',
          title: 'Japanese',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 3,
          lang: 'ja-en',
          title: 'Japanese from English',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 4,
          lang: 'ja-ja',
          title: 'Japanese from Japanese',
          external: true,
          'external-filename': null,
        },
      ];
      if (downloadedPrimaryAdded) {
        tracks.push({
          type: 'sub',
          id: 9,
          lang: 'ja',
          title: path.basename(downloadedPrimaryPath),
          external: true,
          'external-filename': downloadedPrimaryPath,
        });
      }
      return tracks;
    },
    refreshCurrentSubtitle: () => {},
    refreshSubtitleSidebarSource: async (sourcePath) => {
      refreshedSidebarSources.push(sourcePath);
    },
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: () => {
      throw new Error('existing manual tracks should not report failure');
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.equal(selectedPrimarySid, 9);
  assert.equal(selectedSecondarySid, 1);
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'sub-add' && command[1] === downloadedPrimaryPath && command[2] === 'select',
    ),
  );
  assert.deepEqual(refreshedSidebarSources, [downloadedPrimaryPath]);
  assert.equal(
    commands.some((command) => command[0] === 'sub-remove'),
    false,
  );
});

test('youtube flow injects downloaded primary subtitles instead of reusing streamed youtube tracks', async () => {
  const commands: Array<Array<string | number>> = [];
  const refreshedSidebarSources: string[] = [];
  let selectedPrimarySid: number | null = null;
  let downloadedPrimaryAdded = false;
  const downloadedPrimaryPath = '/tmp/subminer-youtube-subtitles-abc/manual-ja.ja.vtt';

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [
        {
          ...primaryTrack,
          id: 'manual:ja',
          sourceLanguage: 'ja',
          kind: 'manual',
          title: 'Japanese',
        },
      ],
    }),
    acquireYoutubeSubtitleTracks: async () => {
      throw new Error('single primary selection should not batch download');
    },
    acquireYoutubeSubtitleTrack: async ({ track }) => {
      assert.equal(track.id, 'manual:ja');
      return { path: downloadedPrimaryPath };
    },
    openPicker: async (payload) => {
      queueMicrotask(() => {
        void runtime.resolveActivePicker({
          sessionId: payload.sessionId,
          action: 'use-selected',
          primaryTrackId: 'manual:ja',
          secondaryTrackId: null,
        });
      });
      return true;
    },
    pauseMpv: () => {},
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      commands.push(command);
      if (
        command[0] === 'sub-add' &&
        command[1] === downloadedPrimaryPath &&
        command[2] === 'select'
      ) {
        downloadedPrimaryAdded = true;
      }
      if (command[0] === 'set_property' && command[1] === 'sid' && typeof command[2] === 'number') {
        selectedPrimarySid = command[2];
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      if (name === 'sid') {
        return selectedPrimarySid;
      }
      return downloadedPrimaryAdded
        ? [
            {
              type: 'sub',
              id: 2,
              lang: 'ja',
              title: 'Japanese',
              external: true,
              'external-filename': '/tmp/mpv-ytdl-track-ja.vtt',
            },
            {
              type: 'sub',
              id: 9,
              lang: 'ja',
              title: path.basename(downloadedPrimaryPath),
              external: true,
              'external-filename': downloadedPrimaryPath,
            },
          ]
        : [
            {
              type: 'sub',
              id: 2,
              lang: 'ja',
              title: 'Japanese',
              external: true,
              'external-filename': '/tmp/mpv-ytdl-track-ja.vtt',
            },
          ];
    },
    refreshCurrentSubtitle: () => {},
    refreshSubtitleSidebarSource: async (sourcePath) => {
      refreshedSidebarSources.push(sourcePath);
    },
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: (message) => {
      throw new Error(message);
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com/watch?v=video123' });

  assert.equal(selectedPrimarySid, 9);
  assert.deepEqual(refreshedSidebarSources, [downloadedPrimaryPath]);
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'sub-add' && command[1] === downloadedPrimaryPath && command[2] === 'select',
    ),
  );
});

test('youtube flow confirms primary subtitle load before sidebar and tokenization waits', async () => {
  const events: string[] = [];
  let selectedPrimarySid: number | null = null;
  let downloadedPrimaryAdded = false;
  const downloadedPrimaryPath = '/tmp/subminer-youtube-subtitles-abc/auto-ja-orig.vtt';

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () => {
      throw new Error('single primary selection should not batch download');
    },
    acquireYoutubeSubtitleTrack: async () => ({ path: downloadedPrimaryPath }),
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
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      if (
        command[0] === 'sub-add' &&
        command[1] === downloadedPrimaryPath &&
        command[2] === 'select'
      ) {
        downloadedPrimaryAdded = true;
      }
      if (command[0] === 'set_property' && command[1] === 'sid' && typeof command[2] === 'number') {
        selectedPrimarySid = command[2];
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      if (name === 'sid') {
        return selectedPrimarySid;
      }
      return downloadedPrimaryAdded
        ? [
            {
              type: 'sub',
              id: 9,
              lang: 'ja-orig',
              title: path.basename(downloadedPrimaryPath),
              external: true,
              'external-filename': downloadedPrimaryPath,
            },
          ]
        : [];
    },
    refreshCurrentSubtitle: () => {},
    refreshSubtitleSidebarSource: async () => {
      events.push('sidebar');
      assert.ok(
        events.includes('notify'),
        'primary load should be confirmed before sidebar parsing can delay',
      );
    },
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {
      events.push('tokenization');
      assert.ok(
        events.includes('notify'),
        'primary load should be confirmed before tokenization waits can delay',
      );
    },
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: (message) => {
      throw new Error(message);
    },
    notifyPrimarySubtitleLoaded: () => {
      events.push('notify');
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com/watch?v=video123' });

  assert.deepEqual(events, ['notify', 'sidebar', 'tokenization']);
});

test('youtube flow downloads subtitles into temporary dirs and exposes cleanup', async () => {
  const outputDirs: string[] = [];
  const cleanupCalls: string[][] = [];
  let tempDirIndex = 0;
  let selectedPrimarySid: number | null = null;
  let addedSubtitlePath: string | null = null;

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () => {
      throw new Error('single primary selection should not batch download');
    },
    acquireYoutubeSubtitleTrack: async ({ outputDir }) => {
      outputDirs.push(outputDir);
      return { path: path.join(outputDir, 'auto-ja-orig.vtt') };
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
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      if (command[0] === 'sub-add' && typeof command[1] === 'string') {
        addedSubtitlePath = command[1];
      }
      if (command[0] === 'set_property' && command[1] === 'sid' && typeof command[2] === 'number') {
        selectedPrimarySid = command[2];
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      if (name === 'sid') {
        return selectedPrimarySid;
      }
      return addedSubtitlePath
        ? [
            {
              type: 'sub',
              id: 10 + tempDirIndex,
              lang: 'ja-orig',
              title: path.basename(addedSubtitlePath),
              external: true,
              'external-filename': addedSubtitlePath,
            },
          ]
        : [];
    },
    refreshCurrentSubtitle: () => {},
    refreshSubtitleSidebarSource: async () => {},
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: (message) => {
      throw new Error(message);
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp/unused-youtube-cache',
    createSubtitleTempDir: async () => {
      tempDirIndex += 1;
      return `/tmp/subminer-youtube-subtitles-${tempDirIndex}`;
    },
    cleanupSubtitleTempDirs: (dirs) => {
      cleanupCalls.push([...dirs]);
    },
  });

  await runtime.openManualPicker({ url: 'https://example.com/watch?v=video123' });
  addedSubtitlePath = null;
  selectedPrimarySid = null;
  await runtime.openManualPicker({ url: 'https://example.com/watch?v=video123' });
  runtime.cleanupSubtitleTempDirs();
  runtime.cleanupSubtitleTempDirs();

  assert.deepEqual(outputDirs, [
    '/tmp/subminer-youtube-subtitles-1',
    '/tmp/subminer-youtube-subtitles-2',
  ]);
  assert.deepEqual(cleanupCalls, [
    ['/tmp/subminer-youtube-subtitles-1'],
    ['/tmp/subminer-youtube-subtitles-2'],
  ]);
});

test('youtube flow falls back to configured output dir when subtitle temp dir creation fails', async () => {
  const outputDirs: string[] = [];
  const warnings: string[] = [];
  let selectedPrimarySid: number | null = null;
  let addedSubtitlePath: string | null = null;

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () => {
      throw new Error('single primary selection should not batch download');
    },
    acquireYoutubeSubtitleTrack: async ({ outputDir }) => {
      outputDirs.push(outputDir);
      return { path: path.join(outputDir, 'auto-ja-orig.vtt') };
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
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      if (command[0] === 'sub-add' && typeof command[1] === 'string') {
        addedSubtitlePath = command[1];
      }
      if (command[0] === 'set_property' && command[1] === 'sid' && typeof command[2] === 'number') {
        selectedPrimarySid = command[2];
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      if (name === 'sid') {
        return selectedPrimarySid;
      }
      return addedSubtitlePath
        ? [
            {
              type: 'sub',
              id: 11,
              lang: 'ja-orig',
              title: path.basename(addedSubtitlePath),
              external: true,
              'external-filename': addedSubtitlePath,
            },
          ]
        : [];
    },
    refreshCurrentSubtitle: () => {},
    refreshSubtitleSidebarSource: async () => {},
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: (message) => {
      throw new Error(message);
    },
    warn: (message) => {
      warnings.push(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp/youtube-cache',
    createSubtitleTempDir: async () => {
      throw new Error('tmp unavailable');
    },
    cleanupSubtitleTempDirs: () => {},
  });

  await runtime.openManualPicker({ url: 'https://example.com/watch?v=video123' });

  assert.deepEqual(outputDirs, ['/tmp/youtube-cache']);
  assert.deepEqual(warnings, [
    'Failed to create YouTube subtitle temp dir; using configured output dir: tmp unavailable',
  ]);
});

test('youtube flow waits for manual secondary tracks while injecting downloaded primary', async () => {
  const commands: Array<Array<string | number>> = [];
  let selectedPrimarySid: number | null = null;
  let selectedSecondarySid: number | null = null;
  let trackListReads = 0;
  let downloadedPrimaryAdded = false;
  const downloadedPrimaryPath = '/tmp/manual-ja.ja.srt';

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [
        {
          ...primaryTrack,
          id: 'manual:ja',
          sourceLanguage: 'ja',
          kind: 'manual',
          title: 'Japanese',
        },
        {
          ...secondaryTrack,
          id: 'manual:en',
          sourceLanguage: 'en',
          kind: 'manual',
          title: 'English',
        },
      ],
    }),
    acquireYoutubeSubtitleTracks: async () => {
      throw new Error('should not batch download when manual tracks appear after startup');
    },
    acquireYoutubeSubtitleTrack: async ({ track }) => {
      if (track.language === 'ja') {
        return { path: downloadedPrimaryPath };
      }
      throw new Error('should not download secondary track when manual english appears in mpv');
    },
    openPicker: async (payload) => {
      queueMicrotask(() => {
        void runtime.resolveActivePicker({
          sessionId: payload.sessionId,
          action: 'use-selected',
          primaryTrackId: 'manual:ja',
          secondaryTrackId: 'manual:en',
        });
      });
      return true;
    },
    pauseMpv: () => {},
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      commands.push(command);
      if (
        command[0] === 'sub-add' &&
        command[1] === downloadedPrimaryPath &&
        command[2] === 'select'
      ) {
        downloadedPrimaryAdded = true;
      }
      if (command[0] === 'set_property' && command[1] === 'sid' && typeof command[2] === 'number') {
        selectedPrimarySid = command[2];
      }
      if (
        command[0] === 'set_property' &&
        command[1] === 'secondary-sid' &&
        typeof command[2] === 'number'
      ) {
        selectedSecondarySid = command[2];
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '字幕です';
      }
      if (name === 'sid') {
        return selectedPrimarySid;
      }
      if (name === 'secondary-sid') {
        return selectedSecondarySid;
      }
      trackListReads += 1;
      if (trackListReads === 1) {
        return [];
      }
      const tracks: Array<Record<string, unknown>> = [
        {
          type: 'sub',
          id: 1,
          lang: 'en',
          title: 'English',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 2,
          lang: 'ja',
          title: 'Japanese',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 3,
          lang: 'ja-en',
          title: 'Japanese from English',
          external: true,
          'external-filename': null,
        },
        {
          type: 'sub',
          id: 4,
          lang: 'ja-ja',
          title: 'Japanese from Japanese',
          external: true,
          'external-filename': null,
        },
      ];
      if (downloadedPrimaryAdded) {
        tracks.push({
          type: 'sub',
          id: 9,
          lang: 'ja',
          title: path.basename(downloadedPrimaryPath),
          external: true,
          'external-filename': downloadedPrimaryPath,
        });
      }
      return tracks;
    },
    refreshCurrentSubtitle: () => {},
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: () => {
      throw new Error('delayed manual tracks should not report failure');
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.openManualPicker({ url: 'https://example.com' });

  assert.equal(selectedPrimarySid, 9);
  assert.equal(selectedSecondarySid, 1);
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'sub-add' && command[1] === downloadedPrimaryPath && command[2] === 'select',
    ),
  );
});

test('youtube flow injects downloaded primary even when reusable manual youtube tracks exist', async () => {
  const commands: Array<Array<string | number>> = [];
  let selectedPrimarySid: number | null = null;
  let selectedSecondarySid: number | null = null;
  let downloadedPrimaryAdded = false;
  const downloadedPrimaryPath = '/tmp/manual-ja.ja.srt';

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [
        {
          id: 'manual:ja',
          language: 'ja',
          sourceLanguage: 'ja',
          kind: 'manual',
          title: 'Japanese',
          label: 'Japanese',
        },
        {
          id: 'manual:en',
          language: 'en',
          sourceLanguage: 'en',
          kind: 'manual',
          title: 'English',
          label: 'English',
        },
      ],
    }),
    acquireYoutubeSubtitleTracks: async () => {
      throw new Error('should not batch-download when existing manual tracks are reusable');
    },
    acquireYoutubeSubtitleTrack: async ({ track }) => {
      if (track.id === 'manual:ja') {
        return { path: downloadedPrimaryPath };
      }
      throw new Error(
        'should not download secondary track when existing manual english track is reusable',
      );
    },
    openPicker: async () => false,
    pauseMpv: () => {},
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      commands.push(command);
      if (
        command[0] === 'sub-add' &&
        command[1] === downloadedPrimaryPath &&
        command[2] === 'select'
      ) {
        downloadedPrimaryAdded = true;
      }
      if (command[0] === 'set_property' && command[1] === 'sid') {
        selectedPrimarySid = Number(command[2]);
      }
      if (command[0] === 'set_property' && command[1] === 'secondary-sid') {
        selectedSecondarySid = Number(command[2]);
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'track-list') {
        const tracks: Array<Record<string, unknown>> = [
          {
            type: 'sub',
            id: 1,
            lang: 'en',
            title: 'English',
            external: true,
            'external-filename': '/tmp/mpv-ytdl-track-en.vtt',
          },
          {
            type: 'sub',
            id: 2,
            lang: 'ja',
            title: 'Japanese',
            external: true,
            'external-filename': '/tmp/mpv-ytdl-track-ja.vtt',
          },
          {
            type: 'sub',
            id: 3,
            lang: 'ja-en',
            title: 'Japanese from English',
            external: true,
            'external-filename': '/tmp/mpv-ytdl-track-ja-en.vtt',
          },
        ];
        if (downloadedPrimaryAdded) {
          tracks.push({
            type: 'sub',
            id: 9,
            lang: 'ja',
            title: path.basename(downloadedPrimaryPath),
            external: true,
            'external-filename': downloadedPrimaryPath,
          });
        }
        return tracks;
      }
      if (name === 'sid') {
        return selectedPrimarySid;
      }
      if (name === 'secondary-sid') {
        return selectedSecondarySid;
      }
      if (name === 'sub-text') {
        return '';
      }
      return null;
    },
    refreshCurrentSubtitle: () => {},
    refreshSubtitleSidebarSource: async () => {},
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: (message) => {
      throw new Error(message);
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.runYoutubePlaybackFlow({
    url: 'https://example.com/watch?v=video123',
    mode: 'download',
  });

  assert.equal(selectedPrimarySid, 9);
  assert.equal(selectedSecondarySid, 1);
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'sub-add' && command[1] === downloadedPrimaryPath && command[2] === 'select',
    ),
  );
});

test('youtube flow falls back to existing auto secondary track when auto secondary download fails', async () => {
  const commands: Array<Array<string | number>> = [];
  let selectedPrimarySid: number | null = null;
  let selectedSecondarySid: number | null = null;
  let primaryTrackAdded = false;

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [
        {
          id: 'auto:ja-orig',
          language: 'ja-orig',
          sourceLanguage: 'ja-orig',
          kind: 'auto',
          title: 'Japanese (Original)',
          label: 'Japanese (Original) (auto)',
        },
        {
          id: 'auto:en',
          language: 'en',
          sourceLanguage: 'en',
          kind: 'auto',
          title: 'English',
          label: 'English (auto)',
        },
      ],
    }),
    acquireYoutubeSubtitleTracks: async () =>
      new Map<string, string>([['auto:ja-orig', '/tmp/auto-ja-orig.ja-orig.vtt']]),
    acquireYoutubeSubtitleTrack: async ({ track }) => {
      if (track.id === 'auto:en') {
        throw new Error('HTTP 429 while downloading en');
      }
      return { path: '/tmp/auto-ja-orig.ja-orig.vtt' };
    },
    openPicker: async () => false,
    pauseMpv: () => {},
    resumeMpv: () => {},
    sendMpvCommand: (command) => {
      commands.push(command);
      if (
        command[0] === 'sub-add' &&
        command[1] === '/tmp/auto-ja-orig.ja-orig.vtt' &&
        command[2] === 'select'
      ) {
        primaryTrackAdded = true;
      }
      if (command[0] === 'set_property' && command[1] === 'sid' && typeof command[2] === 'number') {
        selectedPrimarySid = command[2];
      }
      if (
        command[0] === 'set_property' &&
        command[1] === 'secondary-sid' &&
        typeof command[2] === 'number'
      ) {
        selectedSecondarySid = command[2];
      }
    },
    requestMpvProperty: async (name) => {
      if (name === 'sub-text') {
        return '';
      }
      if (name === 'sid') {
        return selectedPrimarySid;
      }
      if (name === 'secondary-sid') {
        return selectedSecondarySid;
      }
      return primaryTrackAdded
        ? [
            {
              type: 'sub',
              id: 1,
              lang: 'en',
              title: 'English',
              external: true,
              'external-filename': '/tmp/mpv-auto-en.vtt',
            },
            {
              type: 'sub',
              id: 3,
              lang: 'ja-orig',
              title: 'Japanese (Original)',
              external: true,
              'external-filename': '/tmp/mpv-auto-ja-orig.vtt',
            },
            {
              type: 'sub',
              id: 4,
              lang: 'ja-orig',
              title: 'auto-ja-orig.ja-orig.vtt',
              external: true,
              'external-filename': '/tmp/auto-ja-orig.ja-orig.vtt',
            },
          ]
        : [
            {
              type: 'sub',
              id: 1,
              lang: 'en',
              title: 'English',
              external: true,
              'external-filename': '/tmp/mpv-auto-en.vtt',
            },
            {
              type: 'sub',
              id: 3,
              lang: 'ja-orig',
              title: 'Japanese (Original)',
              external: true,
              'external-filename': '/tmp/mpv-auto-ja-orig.vtt',
            },
          ];
    },
    refreshCurrentSubtitle: () => {},
    refreshSubtitleSidebarSource: async () => {},
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async () => {},
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {},
    showMpvOsd: () => {},
    reportSubtitleFailure: (message) => {
      throw new Error(message);
    },
    warn: (message) => {
      throw new Error(message);
    },
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
  });

  await runtime.runYoutubePlaybackFlow({
    url: 'https://example.com/watch?v=video123',
    mode: 'download',
  });

  assert.equal(selectedPrimarySid, 4);
  assert.equal(selectedSecondarySid, 1);
});
