import assert from 'node:assert/strict';
import test from 'node:test';
import type { MediaTimingReviewOpenPayload } from '../../types/anki';
import type { MediaTimingFrameOptions } from '../../core/services/media-timing-frame';
import {
  createMediaTimingReviewRuntime,
  type MediaTimingReviewRuntimeDeps,
} from './media-timing-review';

async function start(
  options: Partial<MediaTimingReviewRuntimeDeps> = {},
  screenshotEnabled = true,
) {
  let opened!: (payload: MediaTimingReviewOpenPayload) => void;
  const payloadPromise = new Promise<MediaTimingReviewOpenPayload>((resolve) => {
    opened = resolve;
  });
  const frames: MediaTimingFrameOptions[] = [];
  const runtime = createMediaTimingReviewRuntime({
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: '/video.mkv',
      send: () => {},
      requestProperty: async (name) => (name === 'duration' ? 100 : true),
    }),
    getCurrentMediaPath: () => '/video.mkv',
    getMpvExecutablePath: () => '',
    createPreviewSession: () => ({
      start: async () => {},
      play: async () => {},
      stop: async () => {},
      onPlaybackEnded: () => {},
      dispose: () => {},
    }),
    generateWaveform: async () => [0, 1],
    resolveVideoSource: async () => ({ path: '/video.mkv' }),
    generateFrame: async (options) => {
      frames.push(options);
      return { dataUrl: 'data:image/jpeg;base64,AA==', timestamp: options.timestamp };
    },
    openModal: async (payload) => {
      opened(payload);
      return true;
    },
    showStatus: () => {},
    ...options,
  });
  const decision = runtime.requestReview({
    kind: 'word',
    text: '字幕',
    startTime: 10,
    endTime: 12,
    audioPadding: 0,
    maxMediaDuration: 30,
    screenshotEnabled,
  });
  return { runtime, decision, payload: await payloadPromise, frames };
}

test('review accepts a screenshot outside the audio range and validates media bounds', async () => {
  const { runtime, decision, payload, frames } = await start();
  for (const timestamp of [NaN, Infinity, -1, 100]) {
    assert.equal((await runtime.getFrame({ reviewId: payload.reviewId, timestamp })).ok, false);
    assert.equal(
      runtime.resolveReview({
        reviewId: payload.reviewId,
        decision: { action: 'confirm', startTime: 10, endTime: 11, screenshotTime: timestamp },
      }).ok,
      false,
    );
  }
  assert.equal((await runtime.getFrame({ reviewId: payload.reviewId, timestamp: 13 })).ok, true);
  assert.equal(frames[0]?.timestamp, 13);
  assert.equal(
    runtime.resolveReview({
      reviewId: payload.reviewId,
      decision: { action: 'confirm', startTime: 10, endTime: 11, screenshotTime: 13 },
    }).ok,
    true,
  );
  assert.deepEqual(await decision, {
    action: 'confirm',
    startTime: 10,
    endTime: 11,
    screenshotTime: 13,
  });
});

test('screenshot preview reuses the remote audio window and its absolute timestamps', async () => {
  let downloads = 0;
  const remote = 'https://example.test/video';
  const { runtime, payload, decision, frames } = await start({
    resolveMediaSource: async () => ({ path: remote }),
    resolveVideoSource: async () => ({ path: remote }),
    acquireMediaWindow: async () => {
      downloads += 1;
      return {
        path: '/window.mkv',
        sourcePath: remote,
        startTime: 7,
        endTime: 15,
        audioStreamIndex: null,
        media: { path: '/window.mkv', absoluteTimestamps: true },
      };
    },
  });
  await runtime.getWaveform({ reviewId: payload.reviewId, startTime: 8, endTime: 14 });
  assert.equal((await runtime.getFrame({ reviewId: payload.reviewId, timestamp: 11 })).ok, true);
  assert.equal(downloads, 1);
  assert.deepEqual(frames[0]?.media, { path: '/window.mkv', absoluteTimestamps: true });
  await runtime.dispose();
  await decision;
});

test('split video streams never read an audio-only cache as a video source', async () => {
  const video = {
    path: 'https://example.test/video',
    inputOptions: { headers: { 'X-Test': 'value' } },
  };
  const { runtime, payload, decision, frames } = await start({
    resolveMediaSource: async () => ({ path: 'https://example.test/audio' }),
    resolveVideoSource: async () => video,
  });
  await runtime.getFrame({ reviewId: payload.reviewId, timestamp: 11 });
  assert.deepEqual(frames[0]?.media, video);
  await runtime.dispose();
  await decision;
});

test('disabled screenshots and missing video inputs do not trigger extraction', async () => {
  for (const [enabled, options] of [
    [false, {}],
    [true, { resolveVideoSource: async () => null }],
  ] as const) {
    const { runtime, payload, decision, frames } = await start(options, enabled);
    assert.equal((await runtime.getFrame({ reviewId: payload.reviewId, timestamp: 11 })).ok, false);
    assert.equal(frames.length, 0);
    await runtime.dispose();
    await decision;
  }
});

test('closing a review invalidates in-flight frame results and clears the frame index', async () => {
  let finish!: (value: { dataUrl: string; timestamp: number }) => void;
  let clearCount = 0;
  const { runtime, payload, decision } = await start({
    generateFrame: () =>
      new Promise((resolve) => {
        finish = resolve;
      }),
    clearFrameCache: () => {
      clearCount += 1;
    },
  });
  const frame = runtime.getFrame({ reviewId: payload.reviewId, timestamp: 11 });
  await runtime.dispose();
  finish({ dataUrl: 'data:image/jpeg;base64,AA==', timestamp: 11 });
  assert.equal((await frame).stale, true);
  await decision;
  assert.equal(clearCount, 1);
});
