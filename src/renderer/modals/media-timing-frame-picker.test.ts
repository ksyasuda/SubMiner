import assert from 'node:assert/strict';
import test from 'node:test';
import { setTimeout as tick } from 'node:timers/promises';
import { createMediaTimingFramePicker } from './media-timing-frame-picker';
import type { MediaTimingReviewFrameRequest, MediaTimingReviewFrameResult } from '../../types/anki';

function fixture() {
  const requests: Array<{
    request: MediaTimingReviewFrameRequest;
    resolve: (result: MediaTimingReviewFrameResult) => void;
  }> = [];
  let stale = false;
  const picker = createMediaTimingFramePicker({
    debounceMs: 0,
    load: (request) => new Promise((resolve) => requests.push({ request, resolve })),
    onChange: () => {},
    onStale: () => {
      stale = true;
      picker.close();
    },
  });
  const complete = async (index: number, result: number | MediaTimingReviewFrameResult) => {
    requests[index]!.resolve(
      typeof result === 'number'
        ? { ok: true, timestamp: result, dataUrl: `data:image/jpeg;base64,${result}` }
        : result,
    );
    await tick(0);
  };
  return { picker, requests, complete, isStale: () => stale };
}

test('a chosen frame stays fixed until Reset restores midpoint tracking', async () => {
  const { picker, requests, complete } = fixture();
  picker.open('r', true, 11);
  await tick(5);
  await complete(0, 11);
  assert.equal(picker.getScreenshotTime(), undefined);
  picker.choose(13);
  await tick(5);
  await complete(1, 13);
  picker.updateMidpoint(12);
  await tick(5);
  assert.equal(picker.getScreenshotTime(), 13);
  assert.equal(requests.length, 2);
  picker.reset();
  await tick(5);
  assert.equal(requests[2]?.request.timestamp, 12);
  await complete(2, 12);
  assert.equal(picker.getScreenshotTime(), undefined);
  picker.updateMidpoint(14);
  await tick(5);
  assert.equal(requests[3]?.request.timestamp, 14);
  await complete(3, 14);
  picker.close();
});

test('scrubbing keeps only the latest requested frame', async () => {
  const { picker, requests, complete } = fixture();
  picker.open('r', true, 11);
  await tick(5);
  picker.choose(12);
  picker.choose(13);
  picker.choose(14);
  await tick(5);
  assert.equal(requests.length, 1);
  assert.equal(picker.getState().blockConfirm, true);
  await complete(0, 11);
  await tick(5);
  assert.equal(picker.getState().timestamp, undefined);
  assert.equal(requests[1]?.request.timestamp, 14);
  await complete(1, 14);
  assert.equal(picker.getScreenshotTime(), 14);
  assert.equal(picker.getState().blockConfirm, false);
  picker.close();
});

test('failed manual previews block confirmation; Reset allows the default fallback', async () => {
  const { picker, complete } = fixture();
  picker.open('r', true, 11);
  await tick(5);
  await complete(0, { ok: false });
  assert.equal(picker.getState().blockConfirm, false);
  picker.choose(12);
  await tick(5);
  await complete(1, { ok: false });
  assert.equal(picker.getState().blockConfirm, true);
  picker.reset();
  assert.equal(picker.getState().blockConfirm, false);
  picker.close();
});

test('replacing a review ignores old frames; stale responses close the current review', async () => {
  const { picker, requests, complete, isStale } = fixture();
  picker.open('old', true, 11);
  await tick(5);
  picker.close();
  picker.open('new', true, 22);
  await tick(5);
  await complete(0, 11);
  await tick(5);
  assert.equal(picker.getState().timestamp, undefined);
  assert.equal(requests[1]?.request.reviewId, 'new');
  await complete(1, { ok: false, stale: true });
  assert.equal(isStale(), true);
  assert.equal(picker.getState().enabled, false);
});

test('disabled screenshots perform no extraction', async () => {
  const { picker, requests } = fixture();
  picker.open('r', false, 11);
  picker.updateMidpoint(12);
  await tick(5);
  assert.equal(requests.length, 0);
  picker.close();
});
