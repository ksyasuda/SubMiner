import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
import type { MediaTimingReviewOpenPayload } from '../../types/anki';
import type { SpeechWaveformOptions } from '../../core/services/media-timing-waveform';
import type {
  RemoteMediaWindow,
  RemoteMediaWindowRange,
  RemoteMediaWindowSource,
} from '../../core/services/remote-media-window-cache';
import type { MediaTimingPreviewSession } from '../../core/services/media-timing-preview';

type MediaTimingPreviewSessionLike = Pick<MediaTimingPreviewSession, 'start'>;
import {
  buildMediaTimingReviewPayload,
  collectMediaTimingContextLines,
  createMediaTimingReviewRuntime,
} from './media-timing-review';

describe('buildMediaTimingReviewPayload', () => {
  test('starts from the padded range and leaves two seconds to drag on each side', () => {
    const payload = buildMediaTimingReviewPayload(
      {
        kind: 'sentence',
        text: '字幕',
        startTime: 10,
        endTime: 12,
        audioPadding: 0.5,
        maxMediaDuration: 30,
      },
      { reviewId: 'review-1', mediaDuration: 100 },
    );

    assert.equal(payload.selectionStartTime, 9.5);
    assert.equal(payload.selectionEndTime, 12.5);
    assert.equal(payload.timelineStartTime, 7.5);
    assert.equal(payload.timelineEndTime, 14.5);
  });

  test('clamps the padded selection and timeline to media bounds', () => {
    const payload = buildMediaTimingReviewPayload(
      {
        kind: 'word',
        text: '字幕',
        startTime: 0.2,
        endTime: 9.8,
        audioPadding: 1,
        maxMediaDuration: 30,
      },
      { reviewId: 'review-2', mediaDuration: 10 },
    );

    assert.equal(payload.selectionStartTime, 0);
    assert.equal(payload.selectionEndTime, 10);
    assert.equal(payload.timelineStartTime, 0);
    assert.equal(payload.timelineEndTime, 10);
  });

  test('keeps an uncapped selection when max media duration is disabled', () => {
    const payload = buildMediaTimingReviewPayload(
      {
        kind: 'sentence',
        text: '字幕',
        startTime: 10,
        endTime: 55,
        audioPadding: 1,
        maxMediaDuration: 0,
      },
      { reviewId: 'review-unlimited', mediaDuration: 100 },
    );

    assert.equal(payload.selectionStartTime, 9);
    assert.equal(payload.selectionEndTime, 56);
    assert.equal(payload.maxMediaDuration, 0);
  });
});

async function startActiveMediaTimingReview(
  options: {
    maxMediaDuration?: number;
    decisionTimeoutMs?: number;
  } = {},
) {
  const previewCalls: Array<[number, number]> = [];
  let publishPayload!: (payload: MediaTimingReviewOpenPayload) => void;
  const openedPayload = new Promise<MediaTimingReviewOpenPayload>((resolve) => {
    publishPayload = resolve;
  });
  const runtime = createMediaTimingReviewRuntime({
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: '/video/show.mkv',
      requestProperty: async (name) => (name === 'duration' ? 100 : name === 'pause' ? true : null),
      send: () => undefined,
    }),
    getCurrentMediaPath: () => '/video/show.mkv',
    getMpvExecutablePath: () => 'mpv',
    generateWaveform: async () => [],
    decisionTimeoutMs: options.decisionTimeoutMs,
    createPreviewSession: () => ({
      start: async () => undefined,
      play: async (startTime, endTime) => {
        previewCalls.push([startTime, endTime]);
      },
      stop: async () => undefined,
      dispose: () => undefined,
    }),
    openModal: async (payload) => {
      publishPayload(payload);
      return true;
    },
    showStatus: () => undefined,
  });
  const pendingDecision = runtime.requestReview({
    kind: 'sentence',
    text: '字幕',
    startTime: 10,
    endTime: 12,
    audioPadding: 0,
    maxMediaDuration: options.maxMediaDuration ?? 30,
  });

  return { runtime, payload: await openedPayload, pendingDecision, previewCalls };
}

test('media timing review pauses playback, resolves exact timing, and restores playing state', async () => {
  const commands: Array<Array<string | number>> = [];
  const previewCalls: Array<[number, number]> = [];
  let runtime: ReturnType<typeof createMediaTimingReviewRuntime>;
  runtime = createMediaTimingReviewRuntime({
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: '/video/show.mkv',
      requestProperty: async (name) =>
        ({ pause: false, duration: 100, aid: 2, volume: 60 })[
          name as 'pause' | 'duration' | 'aid' | 'volume'
        ],
      send: ({ command }) => commands.push(command),
    }),
    getCurrentMediaPath: () => '/video/show.mkv',
    getMpvExecutablePath: () => 'mpv',
    generateWaveform: async () => [],
    createPreviewSession: () => ({
      start: async () => undefined,
      play: async (startTime, endTime) => {
        previewCalls.push([startTime, endTime]);
      },
      stop: async () => undefined,
      dispose: () => undefined,
    }),
    openModal: async (payload) => {
      queueMicrotask(() => {
        void runtime
          .previewRange({
            reviewId: payload.reviewId,
            startTime: 9.5,
            endTime: 12.5,
          })
          .then(() => {
            runtime.resolveReview({
              reviewId: payload.reviewId,
              decision: { action: 'confirm', startTime: 9.5, endTime: 12.5 },
            });
          });
      });
      return true;
    },
    showStatus: () => undefined,
  });

  const decision = await runtime.requestReview({
    kind: 'word',
    text: '字幕',
    startTime: 10,
    endTime: 12,
    noteId: 42,
    audioPadding: 0.5,
    maxMediaDuration: 30,
  });

  assert.deepEqual(decision, { action: 'confirm', startTime: 9.5, endTime: 12.5 });
  assert.deepEqual(commands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'pause', 'no'],
  ]);
  assert.deepEqual(previewCalls, [[9.5, 12.5]]);
});

test('media timing review analyzes the visible range on the selected audio stream', async () => {
  const waveformCalls: SpeechWaveformOptions[] = [];
  let runtime: ReturnType<typeof createMediaTimingReviewRuntime>;
  runtime = createMediaTimingReviewRuntime({
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: '/video/show.mkv',
      currentAudioStreamIndex: 4,
      requestProperty: async (name) => (name === 'duration' ? 100 : name === 'pause' ? true : null),
      send: () => undefined,
    }),
    getCurrentMediaPath: () => '/video/show.mkv',
    getMpvExecutablePath: () => 'mpv',
    generateWaveform: async (options) => {
      waveformCalls.push(options);
      return [0.1, 0.8, 0.2];
    },
    createPreviewSession: () => ({
      start: async () => undefined,
      play: async () => undefined,
      stop: async () => undefined,
      dispose: () => undefined,
    }),
    openModal: async (payload) => {
      const waveform = await runtime.getWaveform({
        reviewId: payload.reviewId,
        startTime: payload.timelineStartTime,
        endTime: payload.timelineEndTime,
      });
      assert.deepEqual(waveform, { ok: true, peaks: [0.1, 0.8, 0.2] });
      runtime.resolveReview({
        reviewId: payload.reviewId,
        decision: { action: 'use-original' },
      });
      return true;
    },
    showStatus: () => undefined,
  });

  await runtime.requestReview({
    kind: 'sentence',
    text: '字幕',
    startTime: 10,
    endTime: 12,
    audioPadding: 0.5,
    maxMediaDuration: 30,
  });

  assert.deepEqual(waveformCalls, [
    {
      mediaPath: '/video/show.mkv',
      startTime: 7.5,
      endTime: 14.5,
      audioStreamIndex: 4,
    },
  ]);
});

const REMOTE_STREAM_URL = 'https://jellyfin.example/Videos/abc/stream?static=true';

function createWindowStub(options: { fail?: boolean } = {}) {
  const calls: Array<{ source: RemoteMediaWindowSource; range: RemoteMediaWindowRange }> = [];
  const acquireMediaWindow = async (
    source: RemoteMediaWindowSource,
    range: RemoteMediaWindowRange,
  ): Promise<RemoteMediaWindow> => {
    calls.push({ source, range });
    if (options.fail) throw new Error('offline');
    const windowPath = `/tmp/window-${range.startTime}-${range.endTime}.mkv`;
    return {
      path: windowPath,
      startTime: range.startTime,
      endTime: range.endTime,
      sourcePath: source.path,
      audioStreamIndex: source.audioStreamIndex ?? null,
      media: {
        path: windowPath,
        source: 'remote-window',
        singleResolvedStream: true,
        absoluteTimestamps: true,
      },
    };
  };
  return { calls, acquireMediaWindow };
}

function createRemoteReviewRuntime(options: {
  windowStub: ReturnType<typeof createWindowStub>;
  waveformCalls: SpeechWaveformOptions[];
  previewStarts: Array<Parameters<MediaTimingPreviewSessionLike['start']>[0]>;
  previewPlays: Array<[string, number, number]>;
  disposed: string[];
  openModal: (
    runtime: ReturnType<typeof createMediaTimingReviewRuntime>,
    payload: MediaTimingReviewOpenPayload,
  ) => Promise<void>;
}) {
  let runtime!: ReturnType<typeof createMediaTimingReviewRuntime>;
  runtime = createMediaTimingReviewRuntime({
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: REMOTE_STREAM_URL,
      currentAudioStreamIndex: 2,
      requestProperty: async (name) =>
        ({ pause: true, duration: 100, aid: 3, volume: 60 })[
          name as 'pause' | 'duration' | 'aid' | 'volume'
        ] ?? null,
      send: () => undefined,
    }),
    getCurrentMediaPath: () => REMOTE_STREAM_URL,
    getMpvExecutablePath: () => 'mpv',
    resolveMediaSource: async () => ({
      path: REMOTE_STREAM_URL,
      inputOptions: { reconnect: true },
    }),
    acquireMediaWindow: options.windowStub.acquireMediaWindow,
    generateWaveform: async (waveformOptions) => {
      options.waveformCalls.push(waveformOptions);
      return [0.1, 0.8, 0.2];
    },
    createPreviewSession: () => {
      let mediaPath = '';
      return {
        start: async (startOptions) => {
          mediaPath = startOptions.mediaPath;
          options.previewStarts.push(startOptions);
        },
        play: async (startTime, endTime) => {
          options.previewPlays.push([mediaPath, startTime, endTime]);
        },
        stop: async () => undefined,
        dispose: () => {
          options.disposed.push(mediaPath);
        },
      };
    },
    openModal: async (payload) => {
      await options.openModal(runtime, payload);
      return true;
    },
    showStatus: () => undefined,
  });
  return runtime;
}

test('media timing review downloads one window of a remote stream for the waveform and preview', async () => {
  const windowStub = createWindowStub();
  const waveformCalls: SpeechWaveformOptions[] = [];
  const previewStarts: Array<Parameters<MediaTimingPreviewSessionLike['start']>[0]> = [];
  const previewPlays: Array<[string, number, number]> = [];
  const disposed: string[] = [];
  const runtime = createRemoteReviewRuntime({
    windowStub,
    waveformCalls,
    previewStarts,
    previewPlays,
    disposed,
    openModal: async (active, payload) => {
      const waveform = await active.getWaveform({
        reviewId: payload.reviewId,
        startTime: payload.timelineStartTime,
        endTime: payload.timelineEndTime,
      });
      assert.deepEqual(waveform, { ok: true, peaks: [0.1, 0.8, 0.2] });
      assert.deepEqual(
        await active.previewRange({ reviewId: payload.reviewId, startTime: 9.5, endTime: 12.5 }),
        { ok: true },
      );
      active.resolveReview({
        reviewId: payload.reviewId,
        decision: { action: 'confirm', startTime: 9.5, endTime: 12.5 },
      });
    },
  });

  const decision = await runtime.requestReview({
    kind: 'word',
    text: '字幕',
    startTime: 10,
    endTime: 12,
    audioPadding: 0.5,
    maxMediaDuration: 30,
  });

  assert.deepEqual(decision, { action: 'confirm', startTime: 9.5, endTime: 12.5 });
  assert.deepEqual(windowStub.calls, [
    {
      source: { path: REMOTE_STREAM_URL, inputOptions: { reconnect: true }, audioStreamIndex: 2 },
      range: { startTime: 7.5, endTime: 14.5 },
    },
  ]);
  assert.deepEqual(waveformCalls, [
    {
      mediaPath: {
        path: '/tmp/window-7.5-14.5.mkv',
        source: 'remote-window',
        singleResolvedStream: true,
        absoluteTimestamps: true,
      },
      startTime: 7.5,
      endTime: 14.5,
    },
  ]);
  assert.deepEqual(previewStarts, [
    {
      mediaPath: '/tmp/window-7.5-14.5.mkv',
      executablePath: 'mpv',
      volume: 60,
      absoluteTimestamps: true,
    },
  ]);
  assert.deepEqual(previewPlays, [['/tmp/window-7.5-14.5.mkv', 9.5, 12.5]]);
  assert.deepEqual(disposed, ['/tmp/window-7.5-14.5.mkv']);
});

test('media timing review restarts the preview on a wider window when the timeline grows', async () => {
  const windowStub = createWindowStub();
  const waveformCalls: SpeechWaveformOptions[] = [];
  const previewStarts: Array<Parameters<MediaTimingPreviewSessionLike['start']>[0]> = [];
  const previewPlays: Array<[string, number, number]> = [];
  const disposed: string[] = [];
  const runtime = createRemoteReviewRuntime({
    windowStub,
    waveformCalls,
    previewStarts,
    previewPlays,
    disposed,
    openModal: async (active, payload) => {
      await active.previewRange({ reviewId: payload.reviewId, startTime: 9.5, endTime: 12.5 });
      // The user revealed two more seconds before the clip.
      await active.getWaveform({ reviewId: payload.reviewId, startTime: 5.5, endTime: 14.5 });
      await active.previewRange({ reviewId: payload.reviewId, startTime: 6, endTime: 12.5 });
      active.resolveReview({ reviewId: payload.reviewId, decision: { action: 'use-original' } });
    },
  });

  await runtime.requestReview({
    kind: 'sentence',
    text: '字幕',
    startTime: 10,
    endTime: 12,
    audioPadding: 0.5,
    maxMediaDuration: 30,
  });

  assert.deepEqual(
    windowStub.calls.map((call) => call.range),
    [
      { startTime: 7.5, endTime: 14.5 },
      { startTime: 5.5, endTime: 14.5 },
    ],
  );
  assert.deepEqual(
    previewStarts.map((start) => start.mediaPath),
    ['/tmp/window-7.5-14.5.mkv', '/tmp/window-5.5-14.5.mkv'],
  );
  assert.deepEqual(previewPlays, [
    ['/tmp/window-7.5-14.5.mkv', 9.5, 12.5],
    ['/tmp/window-5.5-14.5.mkv', 6, 12.5],
  ]);
  assert.deepEqual(disposed, ['/tmp/window-7.5-14.5.mkv', '/tmp/window-5.5-14.5.mkv']);
  assert.equal(waveformCalls[0]?.startTime, 5.5);
});

test('media timing review falls back to the remote stream after one failed window download', async () => {
  const windowStub = createWindowStub({ fail: true });
  const waveformCalls: SpeechWaveformOptions[] = [];
  const previewStarts: Array<Parameters<MediaTimingPreviewSessionLike['start']>[0]> = [];
  const previewPlays: Array<[string, number, number]> = [];
  const disposed: string[] = [];
  const runtime = createRemoteReviewRuntime({
    windowStub,
    waveformCalls,
    previewStarts,
    previewPlays,
    disposed,
    openModal: async (active, payload) => {
      await active.getWaveform({
        reviewId: payload.reviewId,
        startTime: payload.timelineStartTime,
        endTime: payload.timelineEndTime,
      });
      await active.previewRange({ reviewId: payload.reviewId, startTime: 9.5, endTime: 12.5 });
      active.resolveReview({ reviewId: payload.reviewId, decision: { action: 'use-original' } });
    },
  });

  await runtime.requestReview({
    kind: 'word',
    text: '字幕',
    startTime: 10,
    endTime: 12,
    audioPadding: 0.5,
    maxMediaDuration: 30,
  });

  assert.equal(windowStub.calls.length, 1);
  assert.deepEqual(waveformCalls, [
    {
      mediaPath: { path: REMOTE_STREAM_URL, inputOptions: { reconnect: true } },
      startTime: 7.5,
      endTime: 14.5,
      audioStreamIndex: 2,
    },
  ]);
  assert.deepEqual(previewStarts, [
    { mediaPath: REMOTE_STREAM_URL, executablePath: 'mpv', volume: 60, audioTrackId: 3 },
  ]);
  assert.deepEqual(previewPlays, [[REMOTE_STREAM_URL, 9.5, 12.5]]);
});

test('media timing review never downloads windows for local media', async () => {
  const windowStub = createWindowStub();
  let runtime!: ReturnType<typeof createMediaTimingReviewRuntime>;
  runtime = createMediaTimingReviewRuntime({
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: '/video/show.mkv',
      requestProperty: async (name) => (name === 'duration' ? 100 : name === 'pause' ? true : null),
      send: () => undefined,
    }),
    getCurrentMediaPath: () => '/video/show.mkv',
    getMpvExecutablePath: () => 'mpv',
    resolveMediaSource: async () => ({ path: '/video/show.mkv' }),
    acquireMediaWindow: windowStub.acquireMediaWindow,
    generateWaveform: async () => [0.1, 0.8, 0.2],
    createPreviewSession: () => ({
      start: async () => undefined,
      play: async () => undefined,
      stop: async () => undefined,
      dispose: () => undefined,
    }),
    openModal: async (payload) => {
      await runtime.getWaveform({ reviewId: payload.reviewId, startTime: 7.5, endTime: 14.5 });
      runtime.resolveReview({ reviewId: payload.reviewId, decision: { action: 'use-original' } });
      return true;
    },
    showStatus: () => undefined,
  });

  await runtime.requestReview({
    kind: 'sentence',
    text: '字幕',
    startTime: 10,
    endTime: 12,
    audioPadding: 0.5,
    maxMediaDuration: 30,
  });

  assert.equal(windowStub.calls.length, 0);
});

test('media timing review rejects stale and out-of-range actions before allowing discard', async () => {
  const { runtime, payload, pendingDecision, previewCalls } = await startActiveMediaTimingReview({
    maxMediaDuration: 3,
  });

  assert.deepEqual(
    await runtime.previewRange({ reviewId: 'stale-review', startTime: 10, endTime: 12 }),
    { ok: false, message: 'This timing review is no longer active.' },
  );
  assert.deepEqual(
    runtime.resolveReview({
      reviewId: 'stale-review',
      decision: { action: 'confirm', startTime: 10, endTime: 12 },
    }),
    { ok: false, message: 'This timing review is no longer active.' },
  );
  assert.deepEqual(
    runtime.resolveReview({
      reviewId: payload.reviewId,
      decision: { action: 'confirm', startTime: 10, endTime: 14 },
    }),
    { ok: false, message: 'The selected timing range is invalid.' },
  );
  assert.deepEqual(
    runtime.resolveReview({
      reviewId: payload.reviewId,
      decision: { action: 'confirm', startTime: 99, endTime: 100.5 },
    }),
    { ok: false, message: 'The selected timing range is invalid.' },
  );
  assert.deepEqual(
    runtime.resolveReview({
      reviewId: payload.reviewId,
      decision: { action: 'confirm', startTime: 10, endTime: 12, text: '   ' },
    }),
    { ok: false, message: 'The combined sentence text is invalid.' },
  );
  assert.deepEqual(
    runtime.resolveReview({ reviewId: payload.reviewId, decision: { action: 'discard' } }),
    { ok: true },
  );
  assert.deepEqual(await pendingDecision, { action: 'discard' });
  assert.deepEqual(previewCalls, []);
});

test('collectMediaTimingContextLines splits cues around the mined range', () => {
  const cues = [
    { text: '一行目', startTime: 0, endTime: 2 },
    { text: '二行目', startTime: 2.5, endTime: 4 },
    { text: '', startTime: 4.2, endTime: 4.4 },
    { text: '採掘行', startTime: 5, endTime: 7 },
    { text: '四行目', startTime: 7.5, endTime: 9 },
    { text: '五行目', startTime: 9.5, endTime: 11 },
  ];

  const context = collectMediaTimingContextLines({ cues, startTime: 5, endTime: 7 });

  assert.deepEqual(context.previous, [
    { text: '一行目', startTime: 0, endTime: 2 },
    { text: '二行目', startTime: 2.5, endTime: 4 },
  ]);
  assert.deepEqual(context.next, [
    { text: '四行目', startTime: 7.5, endTime: 9 },
    { text: '五行目', startTime: 9.5, endTime: 11 },
  ]);
});

test('collectMediaTimingContextLines falls back to played history when no cues are loaded', () => {
  const context = collectMediaTimingContextLines({
    cues: [],
    fallbackPrevious: [
      { displayText: '前の行', startTime: 1, endTime: 2 },
      { displayText: '採掘行', startTime: 5, endTime: 7 },
    ],
    startTime: 5,
    endTime: 7,
  });

  assert.deepEqual(context.previous, [{ text: '前の行', startTime: 1, endTime: 2 }]);
  assert.deepEqual(context.next, []);
});

test('media timing review watchdog falls back when the renderer stops responding', async () => {
  const { pendingDecision } = await startActiveMediaTimingReview({ decisionTimeoutMs: 0 });

  assert.deepEqual(await pendingDecision, { action: 'use-original' });
});

test('media timing review does not resume playback when the prior state is unavailable', async () => {
  const commands: Array<Array<string | number>> = [];
  let runtime: ReturnType<typeof createMediaTimingReviewRuntime>;
  runtime = createMediaTimingReviewRuntime({
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: '/video/show.mkv',
      requestProperty: async () => null,
      send: ({ command }) => commands.push(command),
    }),
    getCurrentMediaPath: () => '/video/show.mkv',
    getMpvExecutablePath: () => '',
    generateWaveform: async () => [],
    createPreviewSession: () => ({
      start: async () => {
        throw new Error('preview unavailable');
      },
      play: async () => undefined,
      stop: async () => undefined,
      dispose: () => undefined,
    }),
    openModal: async (payload) => {
      queueMicrotask(() => {
        runtime.resolveReview({
          reviewId: payload.reviewId,
          decision: { action: 'use-original' },
        });
      });
      return true;
    },
    showStatus: () => undefined,
  });

  assert.deepEqual(
    await runtime.requestReview({
      kind: 'sentence',
      text: '字幕',
      startTime: 10,
      endTime: 12,
      audioPadding: 0,
      maxMediaDuration: 30,
    }),
    { action: 'use-original' },
  );
  assert.deepEqual(commands, [['set_property', 'pause', 'yes']]);
});

test('media timing review restores playback when setup fails after pausing', async () => {
  const commands: Array<Array<string | number>> = [];
  const runtime = createMediaTimingReviewRuntime({
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: '/video/show.mkv',
      requestProperty: async (name) => (name === 'pause' ? false : null),
      send: ({ command }) => commands.push(command),
    }),
    getCurrentMediaPath: () => '/video/show.mkv',
    getMpvExecutablePath: () => {
      throw new Error('preview setup failed');
    },
    generateWaveform: async () => [],
    createPreviewSession: () => ({
      start: async () => undefined,
      play: async () => undefined,
      stop: async () => undefined,
      dispose: () => undefined,
    }),
    openModal: async () => true,
    showStatus: () => undefined,
  });

  assert.deepEqual(
    await runtime.requestReview({
      kind: 'word',
      text: '字幕',
      startTime: 10,
      endTime: 12,
      audioPadding: 0,
      maxMediaDuration: 30,
    }),
    { action: 'use-original' },
  );
  assert.deepEqual(commands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'pause', 'no'],
  ]);
});

test('disposing an open review settles it with original timing and restores playback', async () => {
  const commands: Array<Array<string | number>> = [];
  const runtime = createMediaTimingReviewRuntime({
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: '/video/show.mkv',
      requestProperty: async (name) => (name === 'pause' ? false : null),
      send: ({ command }) => commands.push(command),
    }),
    getCurrentMediaPath: () => '/video/show.mkv',
    getMpvExecutablePath: () => 'mpv',
    generateWaveform: async () => [],
    createPreviewSession: () => ({
      start: async () => undefined,
      play: async () => undefined,
      stop: async () => undefined,
      dispose: () => undefined,
    }),
    openModal: async () => true,
    showStatus: () => undefined,
  });

  const pending = runtime.requestReview({
    kind: 'word',
    text: '字幕',
    startTime: 10,
    endTime: 12,
    audioPadding: 0,
    maxMediaDuration: 30,
  });
  await new Promise<void>((resolve) => setImmediate(resolve));
  await runtime.dispose();

  assert.deepEqual(await pending, { action: 'use-original' });
  assert.deepEqual(commands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'pause', 'no'],
  ]);
});
