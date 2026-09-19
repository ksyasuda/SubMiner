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
  const complete = (index: number, timestamp: number) =>
    requests[index]!.resolve({
      ok: true,
      timestamp,
      dataUrl: `data:image/jpeg;base64,${timestamp}`,
    });
  return { picker, requests, complete, isStale: () => stale };
}

test('manual image selection stays fixed across audio edits and reset follows midpoint again', async () => {
  const { picker, requests, complete } = fixture();
  picker.open('r', true, 11);
  await tick(5);
  complete(0, 11);
  await tick(0);
  assert.equal(picker.getScreenshotTime(), undefined);
  picker.choose(13);
  await tick(5);
  complete(1, 13);
  await tick(0);
  picker.updateMidpoint(12);
  await tick(5);
  assert.equal(picker.getScreenshotTime(), 13);
  assert.equal(requests.length, 2);
  picker.reset();
  await tick(5);
  assert.equal(requests[2]?.request.timestamp, 12);
  assert.equal(picker.getScreenshotTime(), undefined);
  complete(2, 12);
  await tick(0);
  picker.close();
});

test('scrubbing coalesces requests and never commits an older preview', async () => {
  const { picker, requests, complete } = fixture();
  picker.open('r', true, 11);
  await tick(5);
  picker.choose(12);
  picker.choose(13);
  picker.choose(14);
  await tick(5);
  assert.equal(requests.length, 1);
  assert.equal(picker.getState().blockConfirm, true);
  complete(0, 11);
  await tick(5);
  assert.equal(picker.getState().timestamp, undefined);
  assert.equal(requests[1]?.request.timestamp, 14);
  complete(1, 14);
  await tick(0);
  assert.equal(picker.getScreenshotTime(), 14);
  assert.equal(picker.getState().blockConfirm, false);
  picker.close();
});

test('failed manual previews block confirmation until reset; failed default previews allow audio review', async () => {
  const { picker, requests } = fixture();
  picker.open('r', true, 11);
  await tick(5);
  requests[0]!.resolve({ ok: false });
  await tick(0);
  assert.equal(picker.getState().blockConfirm, false);
  picker.choose(12);
  await tick(5);
  requests[1]!.resolve({ ok: false });
  await tick(0);
  assert.equal(picker.getState().blockConfirm, true);
  picker.reset();
  assert.equal(picker.getState().blockConfirm, false);
  picker.close();
});

test('closing or replacing a review invalidates pending images, and stale responses close the active review', async () => {
  const { picker, requests, complete, isStale } = fixture();
  picker.open('old', true, 11);
  await tick(5);
  picker.close();
  picker.open('new', true, 22);
  await tick(5);
  complete(0, 11);
  await tick(5);
  assert.equal(picker.getState().timestamp, undefined);
  assert.equal(requests[1]?.request.reviewId, 'new');
  requests[1]!.resolve({ ok: false, stale: true });
  await tick(0);
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
