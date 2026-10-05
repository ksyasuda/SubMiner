import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';
import { createYoutubeFlowRuntime } from './youtube-flow';
import type { YoutubeTrackProbeResult } from '../../core/services/youtube/track-probe';
import type { YoutubePickerOpenPayload, YoutubeTrackOption } from '../../types';

type YoutubeFlowDeps = Parameters<typeof createYoutubeFlowRuntime>[0];
type MpvTrack = {
  type: 'sub';
  id: number;
  lang: string;
  title: string;
  external: true;
  'external-filename': string | null;
};
type MpvCommand = Array<string | number>;

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

const manualJa: YoutubeTrackOption = {
  id: 'manual:ja',
  language: 'ja',
  sourceLanguage: 'ja',
  kind: 'manual',
  title: 'Japanese',
  label: 'Japanese',
};

const manualEn: YoutubeTrackOption = {
  id: 'manual:en',
  language: 'en',
  sourceLanguage: 'en',
  kind: 'manual',
  title: 'English',
  label: 'English',
};

function mpvTrack(
  id: number,
  lang: string,
  title: string,
  externalFilename: string | null = null,
): MpvTrack {
  return { type: 'sub', id, lang, title, external: true, 'external-filename': externalFilename };
}

// Tracks mpv's ytdl hook adds for a YouTube URL; the translated ones must never be reused.
const streamedYoutubeTracks = (): MpvTrack[] => [
  mpvTrack(1, 'en', 'English'),
  mpvTrack(2, 'ja', 'Japanese'),
  mpvTrack(3, 'ja-en', 'Japanese from English'),
  mpvTrack(4, 'ja-ja', 'Japanese from Japanese'),
];

/** Where the default fake downloader writes a track. */
function downloadPath(trackId: string, outputDir = '/tmp'): string {
  return path.join(outputDir, `${trackId.replace(/[^a-z0-9_-]+/gi, '-')}.vtt`);
}

type FakeMpvOptions = {
  tracks?: MpvTrack[];
  subText?: string;
  /** When false, tracks added with sub-add only expose their file name through the title. */
  reportAddedFilenames?: boolean;
  /** Intercepts property reads; `read` returns what the fake mpv would report. */
  readProperty?: (name: string, read: () => unknown) => unknown;
};

/**
 * A small stateful mpv: sub-add appends a track, set_property updates sid/secondary-sid, and
 * reads report the current state.
 */
function createFakeMpv(options: FakeMpvOptions, timeline: string[]) {
  const tracks = [...(options.tracks ?? [])];
  const state = { sid: null as number | null, secondarySid: null as number | null };
  const commands: MpvCommand[] = [];
  let nextId = Math.max(8, ...tracks.map((track) => track.id)) + 1;

  const read = (name: string): unknown => {
    switch (name) {
      case 'track-list':
        return tracks.map((track) => ({ ...track }));
      case 'sid':
        return state.sid;
      case 'secondary-sid':
        return state.secondarySid;
      case 'sub-text':
        return options.subText ?? '字幕です';
      default:
        return null;
    }
  };

  return {
    commands,
    state,
    tracks,
    selectedTrack: (property: 'sid' | 'secondary-sid'): MpvTrack | null => {
      const id = property === 'sid' ? state.sid : state.secondarySid;
      return tracks.find((track) => track.id === id) ?? null;
    },
    sendMpvCommand: (command: MpvCommand): void => {
      commands.push(command);
      const [name, arg1, arg2] = command;
      if (name === 'script-message') timeline.push(String(arg1));
      if (name === 'sub-add') {
        timeline.push('sub-add');
        const filePath = String(arg1);
        tracks.push(
          mpvTrack(
            nextId++,
            String(command[4] ?? ''),
            path.basename(filePath),
            options.reportAddedFilenames === false ? null : filePath,
          ),
        );
      }
      if (name === 'set_property' && (arg1 === 'sid' || arg1 === 'secondary-sid')) {
        const id = typeof arg2 === 'number' ? arg2 : null;
        if (arg1 === 'sid') state.sid = id;
        else state.secondarySid = id;
      }
    },
    requestMpvProperty: async (name: string): Promise<unknown> =>
      options.readProperty ? options.readProperty(name, () => read(name)) : read(name),
  };
}

type FlowHarnessOptions = {
  probeTracks?: YoutubeTrackOption[];
  /** Picker answer; omitted means the user continues without subtitles. */
  pick?: { primaryTrackId: string | null; secondaryTrackId: string | null };
  mpv?: FakeMpvOptions;
  deps?: Partial<YoutubeFlowDeps>;
};

function createFlowHarness(options: FlowHarnessOptions = {}) {
  const timeline: string[] = [];
  const mpv = createFakeMpv(options.mpv ?? {}, timeline);
  const recorded = {
    osd: [] as string[],
    warnings: [] as string[],
    failures: [] as string[],
    waits: [] as number[],
    sidebarSources: [] as string[],
    refreshedSubtitles: [] as string[],
    openedPayloads: [] as YoutubePickerOpenPayload[],
    singleDownloads: [] as string[],
    batchDownloads: [] as string[][],
    counts: { loadedSignals: 0, focus: 0 },
  };

  const runtime: ReturnType<typeof createYoutubeFlowRuntime> = createYoutubeFlowRuntime({
    probeYoutubeTracks: async () => ({
      videoId: 'video123',
      title: 'Video 123',
      tracks: options.probeTracks ?? [primaryTrack],
    }),
    acquireYoutubeSubtitleTrack: async ({ track, outputDir }) => {
      recorded.singleDownloads.push(track.id);
      return { path: downloadPath(track.id, outputDir) };
    },
    acquireYoutubeSubtitleTracks: async ({ tracks, outputDir }) => {
      recorded.batchDownloads.push(tracks.map((track) => track.id));
      return new Map(tracks.map((track) => [track.id, downloadPath(track.id, outputDir)]));
    },
    openPicker: async (payload) => {
      recorded.openedPayloads.push(payload);
      queueMicrotask(() => {
        const { sessionId } = payload;
        void runtime.resolveActivePicker(
          options.pick
            ? { sessionId, action: 'use-selected', ...options.pick }
            : {
                sessionId,
                action: 'continue-without-subtitles',
                primaryTrackId: null,
                secondaryTrackId: null,
              },
        );
      });
      return true;
    },
    pauseMpv: () => timeline.push('pause'),
    resumeMpv: () => timeline.push('resume'),
    sendMpvCommand: mpv.sendMpvCommand,
    requestMpvProperty: mpv.requestMpvProperty,
    refreshCurrentSubtitle: (text) => recorded.refreshedSubtitles.push(text),
    refreshSubtitleSidebarSource: async (sourcePath) => {
      recorded.sidebarSources.push(sourcePath);
    },
    startTokenizationWarmups: async () => {},
    waitForTokenizationReady: async () => {},
    waitForAnkiReady: async () => {},
    wait: async (ms) => {
      recorded.waits.push(ms);
    },
    waitForPlaybackWindowReady: async () => {},
    waitForOverlayGeometryReady: async () => {},
    focusOverlayWindow: () => {
      recorded.counts.focus += 1;
    },
    showMpvOsd: (text) => recorded.osd.push(text),
    reportSubtitleFailure: (message) => recorded.failures.push(message),
    notifyPrimarySubtitleLoaded: () => {
      recorded.counts.loadedSignals += 1;
    },
    warn: (message) => recorded.warnings.push(message),
    log: () => {},
    getYoutubeOutputDir: () => '/tmp',
    ...options.deps,
  });

  return { runtime, mpv, timeline, ...recorded };
}

type FlowHarness = ReturnType<typeof createFlowHarness>;

function assertNoProblems(harness: FlowHarness): void {
  assert.deepEqual(harness.warnings, []);
  assert.deepEqual(harness.failures, []);
}

const PRIMARY_FAILURE =
  'Primary subtitles failed to load. Use the YouTube subtitle picker to try manually.';

test('youtube flow announces manual picker opening before probing tracks', async () => {
  let resolveProbe: (probe: YoutubeTrackProbeResult) => void = () => {};
  const probePromise = new Promise<YoutubeTrackProbeResult>((resolve) => {
    resolveProbe = resolve;
  });
  const harness = createFlowHarness({ deps: { probeYoutubeTracks: () => probePromise } });

  const pending = harness.runtime.openManualPicker({ url: 'https://example.com' });
  await Promise.resolve();

  assert.deepEqual(harness.osd, ['Opening YouTube subtitle picker...']);

  resolveProbe({ videoId: 'video123', title: 'Video 123', tracks: [] });
  await pending;
  assertNoProblems(harness);
});

test('youtube flow can open a manual picker session and load the selected subtitles', async () => {
  const harness = createFlowHarness({
    probeTracks: [primaryTrack, secondaryTrack],
    pick: { primaryTrackId: primaryTrack.id, secondaryTrackId: secondaryTrack.id },
  });

  await harness.runtime.openManualPicker({ url: 'https://example.com' });

  assert.equal(harness.openedPayloads.length, 1);
  assert.equal(harness.openedPayloads[0]?.defaultPrimaryTrackId, primaryTrack.id);
  assert.equal(harness.openedPayloads[0]?.defaultSecondaryTrackId, secondaryTrack.id);
  assert.ok(harness.waits.includes(150));
  assert.deepEqual(harness.batchDownloads, [[primaryTrack.id, secondaryTrack.id]]);
  assert.deepEqual(harness.osd, [
    'Opening YouTube subtitle picker...',
    'Getting subtitles...',
    'Downloading subtitles...',
    'Loading subtitles...',
    'Primary and secondary subtitles loaded.',
  ]);

  const primaryPath = downloadPath(primaryTrack.id);
  assert.equal(harness.mpv.selectedTrack('sid')?.['external-filename'], primaryPath);
  assert.equal(
    harness.mpv.selectedTrack('secondary-sid')?.['external-filename'],
    downloadPath(secondaryTrack.id),
  );
  const visibility = (property: string) =>
    harness.mpv.commands.filter(
      (command) => command[0] === 'set_property' && command[1] === property && command[2] === 'yes',
    ).length;
  assert.equal(visibility('sub-visibility'), 1);
  assert.equal(visibility('secondary-sub-visibility'), 0);
  assert.deepEqual(harness.refreshedSubtitles, ['字幕です']);
  assert.deepEqual(harness.sidebarSources, [primaryPath]);
  assert.equal(harness.counts.focus, 1);
  assertNoProblems(harness);
});

test('youtube flow retries secondary after partial batch subtitle failure', async () => {
  const harness = createFlowHarness({
    probeTracks: [primaryTrack, secondaryTrack],
    pick: { primaryTrackId: primaryTrack.id, secondaryTrackId: secondaryTrack.id },
    deps: {
      acquireYoutubeSubtitleTracks: async () =>
        new Map([[primaryTrack.id, downloadPath(primaryTrack.id)]]),
    },
  });

  await harness.runtime.openManualPicker({ url: 'https://example.com' });

  assert.deepEqual(harness.singleDownloads, [secondaryTrack.id]);
  assert.ok(harness.waits.includes(350));
  assert.equal(
    harness.mpv.selectedTrack('secondary-sid')?.['external-filename'],
    downloadPath(secondaryTrack.id),
  );
  assertNoProblems(harness);
});

test('youtube flow reports probe failure through the configured reporter in manual mode', async () => {
  const harness = createFlowHarness({
    deps: {
      probeYoutubeTracks: async () => {
        throw new Error('probe failed');
      },
    },
  });

  await harness.runtime.openManualPicker({ url: 'https://example.com' });

  assert.deepEqual(harness.failures, [PRIMARY_FAILURE]);
  assert.equal(harness.counts.focus, 1);
});

const missingSubTextCases: Array<{ name: string; mpv: FakeMpvOptions }> = [
  { name: 'before cue text appears', mpv: { subText: '' } },
  {
    name: 'when mpv reports sub-text as unavailable',
    mpv: {
      readProperty: (name, read) => {
        if (name === 'sub-text') {
          throw new Error("Failed to read MPV property 'sub-text': property unavailable");
        }
        return read();
      },
    },
  },
];

for (const missingSubText of missingSubTextCases) {
  test(`youtube flow treats a bound track as loaded ${missingSubText.name}`, async () => {
    const harness = createFlowHarness({
      pick: { primaryTrackId: primaryTrack.id, secondaryTrackId: null },
      mpv: missingSubText.mpv,
    });

    await harness.runtime.openManualPicker({ url: 'https://example.com' });

    assert.equal(harness.counts.loadedSignals, 1);
    assert.deepEqual(harness.refreshedSubtitles, []);
    assertNoProblems(harness);
  });
}

test('youtube flow binds tracks by title and retries secondary selection until mpv reports it', async () => {
  let secondarySidReads = 0;
  const harness = createFlowHarness({
    probeTracks: [primaryTrack, secondaryTrack],
    pick: { primaryTrackId: primaryTrack.id, secondaryTrackId: secondaryTrack.id },
    mpv: {
      reportAddedFilenames: false,
      // mpv ignores the first secondary-sid selection.
      readProperty: (name, read) => {
        if (name !== 'secondary-sid') return read();
        secondarySidReads += 1;
        return secondarySidReads >= 2 ? read() : null;
      },
    },
  });

  await harness.runtime.openManualPicker({ url: 'https://example.com' });

  const secondary = harness.mpv.selectedTrack('secondary-sid');
  assert.equal(secondary?.title, path.basename(downloadPath(secondaryTrack.id)));
  assert.equal(
    harness.mpv.commands.filter(
      (command) =>
        command[0] === 'set_property' &&
        command[1] === 'secondary-sid' &&
        command[2] === secondary?.id,
    ).length,
    2,
  );
  assert.ok(harness.waits.includes(100));
  assertNoProblems(harness);
});

test('youtube flow reuses the matching existing manual secondary track instead of a loose language match', async () => {
  const titledSecondary = { ...secondaryTrack, title: 'manual-en.vtt' };
  const harness = createFlowHarness({
    probeTracks: [primaryTrack, titledSecondary],
    pick: { primaryTrackId: primaryTrack.id, secondaryTrackId: titledSecondary.id },
    mpv: { tracks: [mpvTrack(6, 'en', 'English'), mpvTrack(8, 'en', 'manual-en.vtt')] },
  });

  await harness.runtime.openManualPicker({ url: 'https://example.com' });

  assert.equal(harness.mpv.state.secondarySid, 8);
  assert.deepEqual(harness.singleDownloads, [primaryTrack.id]);
  assertNoProblems(harness);
});

const reuseCases: Array<{
  name: string;
  probeTracks: YoutubeTrackOption[];
  /** Null runs the automatic flow instead of the manual picker. */
  pick: FlowHarnessOptions['pick'] | null;
  mpv: () => FakeMpvOptions;
  expectedSecondarySid: number | null;
}> = [
  {
    name: 'while reusing existing manual secondary tracks',
    probeTracks: [manualJa, manualEn],
    pick: { primaryTrackId: manualJa.id, secondaryTrackId: manualEn.id },
    mpv: () => ({ tracks: streamedYoutubeTracks() }),
    expectedSecondarySid: 1,
  },
  {
    name: 'while waiting for manual secondary tracks that appear after startup',
    probeTracks: [manualJa, manualEn],
    pick: { primaryTrackId: manualJa.id, secondaryTrackId: manualEn.id },
    mpv: () => {
      let trackListReads = 0;
      return {
        tracks: streamedYoutubeTracks(),
        readProperty: (name, read) =>
          name === 'track-list' && ++trackListReads === 1 ? [] : read(),
      };
    },
    expectedSecondarySid: 1,
  },
  {
    name: 'even when reusable manual youtube tracks exist in the automatic flow',
    probeTracks: [manualJa, manualEn],
    pick: null,
    mpv: () => ({
      tracks: [
        mpvTrack(1, 'en', 'English', '/tmp/mpv-ytdl-track-en.vtt'),
        mpvTrack(2, 'ja', 'Japanese', '/tmp/mpv-ytdl-track-ja.vtt'),
        mpvTrack(3, 'ja-en', 'Japanese from English', '/tmp/mpv-ytdl-track-ja-en.vtt'),
      ],
    }),
    expectedSecondarySid: 1,
  },
  {
    name: 'instead of reusing streamed youtube tracks',
    probeTracks: [manualJa],
    pick: { primaryTrackId: manualJa.id, secondaryTrackId: null },
    mpv: () => ({ tracks: [mpvTrack(2, 'ja', 'Japanese', '/tmp/mpv-ytdl-track-ja.vtt')] }),
    expectedSecondarySid: null,
  },
  {
    name: 'and leaves non-authoritative youtube tracks in place',
    probeTracks: [primaryTrack, secondaryTrack],
    pick: { primaryTrackId: primaryTrack.id, secondaryTrackId: secondaryTrack.id },
    mpv: () => ({ tracks: streamedYoutubeTracks() }),
    expectedSecondarySid: 1,
  },
];

for (const reuse of reuseCases) {
  test(`youtube flow injects downloaded primary ${reuse.name}`, async () => {
    const harness = createFlowHarness({
      probeTracks: reuse.probeTracks,
      pick: reuse.pick ?? undefined,
      mpv: reuse.mpv(),
    });
    const url = 'https://example.com/watch?v=video123';

    if (reuse.pick) {
      await harness.runtime.openManualPicker({ url });
    } else {
      await harness.runtime.runYoutubePlaybackFlow({ url });
    }

    const primaryId = reuse.probeTracks[0]!.id;
    const primaryPath = downloadPath(primaryId);
    assert.equal(harness.mpv.selectedTrack('sid')?.['external-filename'], primaryPath);
    assert.equal(harness.mpv.state.secondarySid, reuse.expectedSecondarySid);
    // Reused secondary tracks are never downloaded again.
    assert.deepEqual(harness.singleDownloads, [primaryId]);
    assert.deepEqual(harness.batchDownloads, []);
    assert.deepEqual(harness.sidebarSources, [primaryPath]);
    assert.equal(
      harness.mpv.commands.some((command) => command[0] === 'sub-remove'),
      false,
    );
    assertNoProblems(harness);
  });
}

test('youtube flow falls back to existing auto secondary track when auto secondary download fails', async () => {
  const autoJa: YoutubeTrackOption = {
    id: 'auto:ja-orig',
    language: 'ja-orig',
    sourceLanguage: 'ja-orig',
    kind: 'auto',
    title: 'Japanese (Original)',
    label: 'Japanese (Original) (auto)',
  };
  const autoEn: YoutubeTrackOption = {
    id: 'auto:en',
    language: 'en',
    sourceLanguage: 'en',
    kind: 'auto',
    title: 'English',
    label: 'English (auto)',
  };
  const primaryPath = downloadPath(autoJa.id);
  const harness = createFlowHarness({
    probeTracks: [autoJa, autoEn],
    mpv: {
      tracks: [
        mpvTrack(1, 'en', 'English', '/tmp/mpv-auto-en.vtt'),
        mpvTrack(3, 'ja-orig', 'Japanese (Original)', '/tmp/mpv-auto-ja-orig.vtt'),
      ],
    },
    deps: {
      acquireYoutubeSubtitleTracks: async () => new Map([[autoJa.id, primaryPath]]),
      acquireYoutubeSubtitleTrack: async ({ track }) => {
        if (track.id === autoEn.id) throw new Error('HTTP 429 while downloading en');
        return { path: primaryPath };
      },
    },
  });

  await harness.runtime.runYoutubePlaybackFlow({ url: 'https://example.com/watch?v=video123' });

  assert.equal(harness.mpv.selectedTrack('sid')?.['external-filename'], primaryPath);
  assert.equal(harness.mpv.state.secondarySid, 1);
  assert.deepEqual(harness.failures, []);
});

test('youtube flow confirms primary subtitle load before sidebar and tokenization waits', async () => {
  const events: string[] = [];
  const harness = createFlowHarness({
    pick: { primaryTrackId: primaryTrack.id, secondaryTrackId: null },
    deps: {
      notifyPrimarySubtitleLoaded: () => events.push('notify'),
      refreshSubtitleSidebarSource: async () => {
        events.push('sidebar');
      },
      waitForTokenizationReady: async () => {
        events.push('tokenization');
      },
    },
  });

  await harness.runtime.openManualPicker({ url: 'https://example.com/watch?v=video123' });

  assert.deepEqual(events, ['notify', 'sidebar', 'tokenization']);
});

test('youtube flow downloads subtitles into temporary dirs and exposes cleanup', async () => {
  const cleanupCalls: string[][] = [];
  const outputDirs: string[] = [];
  let tempDirIndex = 0;
  const harness = createFlowHarness({
    pick: { primaryTrackId: primaryTrack.id, secondaryTrackId: null },
    deps: {
      getYoutubeOutputDir: () => '/tmp/unused-youtube-cache',
      acquireYoutubeSubtitleTrack: async ({ track, outputDir }) => {
        outputDirs.push(outputDir);
        return { path: downloadPath(track.id, outputDir) };
      },
      createSubtitleTempDir: async () => {
        tempDirIndex += 1;
        return `/tmp/subminer-youtube-subtitles-${tempDirIndex}`;
      },
      cleanupSubtitleTempDirs: (dirs) => {
        cleanupCalls.push([...dirs]);
      },
    },
  });

  await harness.runtime.openManualPicker({ url: 'https://example.com/watch?v=video123' });
  await harness.runtime.openManualPicker({ url: 'https://example.com/watch?v=video123' });
  harness.runtime.cleanupSubtitleTempDirs();
  harness.runtime.cleanupSubtitleTempDirs();

  assert.deepEqual(outputDirs, [
    '/tmp/subminer-youtube-subtitles-1',
    '/tmp/subminer-youtube-subtitles-2',
  ]);
  // Each new load cleans the previous dir; the explicit cleanup removes the last one once.
  assert.deepEqual(cleanupCalls, [
    ['/tmp/subminer-youtube-subtitles-1'],
    ['/tmp/subminer-youtube-subtitles-2'],
  ]);
  assertNoProblems(harness);
});

test('youtube flow falls back to configured output dir when subtitle temp dir creation fails', async () => {
  const outputDirs: string[] = [];
  const harness = createFlowHarness({
    pick: { primaryTrackId: primaryTrack.id, secondaryTrackId: null },
    deps: {
      getYoutubeOutputDir: () => '/tmp/youtube-cache',
      acquireYoutubeSubtitleTrack: async ({ track, outputDir }) => {
        outputDirs.push(outputDir);
        return { path: downloadPath(track.id, outputDir) };
      },
      createSubtitleTempDir: async () => {
        throw new Error('tmp unavailable');
      },
      cleanupSubtitleTempDirs: () => {},
    },
  });

  await harness.runtime.openManualPicker({ url: 'https://example.com/watch?v=video123' });

  assert.deepEqual(outputDirs, ['/tmp/youtube-cache']);
  assert.deepEqual(harness.warnings, [
    'Failed to create YouTube subtitle temp dir; using configured output dir: tmp unavailable',
  ]);
  assert.deepEqual(harness.failures, []);
});

const WHISPER_URL = 'https://www.youtube.com/watch?v=abcdefghijk';

function createWhisperFlowHarness(
  generate: NonNullable<YoutubeFlowDeps['generateWhisperSubtitles']>,
) {
  let modalOpened = false;
  const mustNotFetchCaptions = async (): Promise<never> => {
    throw new Error('Whisper mode must not probe or download YouTube captions');
  };
  const harness = createFlowHarness({
    deps: {
      probeYoutubeTracks: mustNotFetchCaptions,
      acquireYoutubeSubtitleTrack: mustNotFetchCaptions,
      acquireYoutubeSubtitleTracks: mustNotFetchCaptions,
      getSubtitleSource: () => 'whisper',
      generateWhisperSubtitles: (input) => {
        harness.timeline.push('generate');
        return generate(input);
      },
      openSubtitleGenerationModal: async () => {
        modalOpened = true;
        return true;
      },
    },
  });
  const addedSubtitlePaths = () =>
    harness.mpv.commands.filter((command) => command[0] === 'sub-add').map((command) => command[1]);
  return { ...harness, modalOpened: () => modalOpened, addedSubtitlePaths };
}

test('whisper subtitle source keeps the video paused until the generated subtitles load', async () => {
  const harness = createWhisperFlowHarness(async (input) => input.outputPath);

  await harness.runtime.runYoutubePlaybackFlow({ url: WHISPER_URL });

  assert.deepEqual(harness.timeline, [
    'pause',
    'subminer-autoplay-hold',
    'generate',
    'sub-add',
    'subminer-autoplay-ready',
    'resume',
  ]);
  assert.equal(harness.modalOpened(), true);
  assert.deepEqual(
    harness.addedSubtitlePaths().map((filePath) => path.basename(String(filePath))),
    ['youtube-whisper.ja.srt'],
  );
  assert.deepEqual(harness.failures, []);
});

test('cancelling whisper generation resumes playback without subtitles or an error', async () => {
  const harness = createWhisperFlowHarness(async () => null);

  await harness.runtime.runYoutubePlaybackFlow({ url: WHISPER_URL });

  assert.deepEqual(harness.addedSubtitlePaths(), []);
  assert.deepEqual(harness.failures, []);
  assert.equal(harness.timeline.at(-1), 'resume');
});

test('whisper generation stops quietly and leaves playback alone when the video changes', async () => {
  let signal: AbortSignal | null = null;
  const harness = createWhisperFlowHarness(
    (input) =>
      new Promise((_, reject) => {
        signal = input.signal;
        input.signal.addEventListener('abort', () => reject(new Error('cancelled')));
      }),
  );

  const flow = harness.runtime.runYoutubePlaybackFlow({ url: WHISPER_URL });
  while (!signal) await new Promise((resolve) => setImmediate(resolve));
  harness.runtime.handleMediaPathChange(WHISPER_URL);
  assert.equal((signal as AbortSignal).aborted, false);
  harness.runtime.handleMediaPathChange('https://www.youtube.com/watch?v=zyxwvutsrqp');
  await flow;

  assert.equal((signal as AbortSignal).aborted, true);
  assert.deepEqual(harness.addedSubtitlePaths(), []);
  assert.deepEqual(harness.failures, []);
  assert.equal(harness.timeline.includes('resume'), false);
});

test('whisper generation failures are reported and playback resumes', async () => {
  const harness = createWhisperFlowHarness(async () => {
    throw new Error('No Whisper model found.');
  });

  await harness.runtime.runYoutubePlaybackFlow({ url: WHISPER_URL });

  assert.deepEqual(harness.failures, ['Whisper subtitles failed: No Whisper model found.']);
  assert.deepEqual(harness.addedSubtitlePaths(), []);
  assert.equal(harness.timeline.at(-1), 'resume');
});
