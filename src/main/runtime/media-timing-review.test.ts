import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
import type { MediaTimingReviewOpenPayload } from '../../types/anki';
import {
  buildMediaTimingReviewPayload,
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
  const waveformCalls: Array<{
    mediaPath: string;
    startTime: number;
    endTime: number;
    audioStreamIndex?: number;
  }> = [];
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
    runtime.resolveReview({ reviewId: payload.reviewId, decision: { action: 'discard' } }),
    { ok: true },
  );
  assert.deepEqual(await pendingDecision, { action: 'discard' });
  assert.deepEqual(previewCalls, []);
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
    getMpvExecutablePath: () => 'mpv',
    generateWaveform: async () => [],
    createPreviewSession: () => {
      throw new Error('preview setup failed');
    },
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
