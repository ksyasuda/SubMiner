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
    acquireYoutubeSubtitleTrack: async ({ track }) => ({ path: `/tmp/${track.id}.vtt` }),
    retimeYoutubePrimaryTrack: async ({ primaryPath }) => `${primaryPath}.retimed`,
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
          'external-filename': '/tmp/auto-ja-orig.vtt.retimed',
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
        command[1] === '/tmp/auto-ja-orig.vtt.retimed' &&
        command[2] === 'select',
    ),
  );
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'set_property' &&
        command[1] === 'sub-visibility' &&
        command[2] === 'yes',
    ),
  );
  assert.ok(
    commands.some(
      (command) =>
        command[0] === 'set_property' &&
        command[1] === 'secondary-sub-visibility' &&
        command[2] === 'yes',
    ),
  );
  assert.deepEqual(refreshedSidebarSources, ['/tmp/auto-ja-orig.vtt.retimed']);
  assert.deepEqual(focusOverlayCalls, ['focus-overlay']);
});

test('youtube flow retries secondary after partial batch subtitle failure', async () => {
  const acquireSingleCalls: string[] = [];
  const commands: Array<Array<string | number>> = [];
  const waits: number[] = [];

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
    retimeYoutubePrimaryTrack: async ({ primaryPath }) => primaryPath,
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
        command[0] === 'sub-add' &&
        command[1] === '/tmp/manual:en.vtt' &&
        command[2] === 'cached',
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
    retimeYoutubePrimaryTrack: async ({ primaryPath }) => primaryPath,
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

  const runtime = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: [primaryTrack],
    }),
    acquireYoutubeSubtitleTracks: async () => new Map(),
    acquireYoutubeSubtitleTrack: async () => ({ path: '/tmp/auto-ja-orig.vtt' }),
    retimeYoutubePrimaryTrack: async ({ primaryPath }) => primaryPath,
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
    acquireYoutubeSubtitleTrack: async ({ track }) => ({ path: `/tmp/${track.id}.vtt` }),
    retimeYoutubePrimaryTrack: async ({ primaryPath }) => primaryPath,
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
          title: 'English',
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
