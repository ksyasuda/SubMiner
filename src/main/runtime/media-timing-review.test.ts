import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
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
