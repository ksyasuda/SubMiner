import assert from 'node:assert/strict';
import test from 'node:test';

import {
  createMaybeStartOverlayLoadingOsdHandler,
  shouldStartOverlayLoadingOsd,
} from './overlay-loading-osd-start';

test('overlay loading OSD starts for visible overlay before content is ready', () => {
  assert.equal(
    shouldStartOverlayLoadingOsd({
      visibleOverlayRequested: true,
      overlayContentReady: false,
    }),
    true,
  );
});

test('overlay loading OSD does not start when hidden or already ready', () => {
  assert.equal(
    shouldStartOverlayLoadingOsd({
      visibleOverlayRequested: false,
      overlayContentReady: false,
    }),
    false,
  );
  assert.equal(
    shouldStartOverlayLoadingOsd({
      visibleOverlayRequested: true,
      overlayContentReady: true,
    }),
    false,
  );
});

test('overlay loading OSD media-path trigger ignores empty paths', () => {
  assert.equal(
    shouldStartOverlayLoadingOsd({
      visibleOverlayRequested: true,
      overlayContentReady: false,
      mediaPath: '   ',
    }),
    false,
  );
});

test('overlay loading OSD handler starts idempotent status through injected deps', () => {
  const calls: string[] = [];
  const maybeStart = createMaybeStartOverlayLoadingOsdHandler({
    getVisibleOverlayRequested: () => true,
    isOverlayContentReady: () => false,
    startOverlayLoadingOsd: () => {
      calls.push('start');
    },
  });

  maybeStart();
  maybeStart('/tmp/video.mkv');
  maybeStart('   ');

  assert.deepEqual(calls, ['start', 'start']);
});
