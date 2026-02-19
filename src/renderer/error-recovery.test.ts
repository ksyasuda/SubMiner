import test from 'node:test';
import assert from 'node:assert/strict';

import { createRendererRecoveryController } from './error-recovery.js';

test('handleError logs context and recovers overlay state', () => {
  const payloads: unknown[] = [];
  let dismissed = 0;
  let restored = 0;
  const shown: string[] = [];

  const controller = createRendererRecoveryController({
    dismissActiveUi: () => {
      dismissed += 1;
    },
    restoreOverlayInteraction: () => {
      restored += 1;
    },
    showToast: (message) => {
      shown.push(message);
    },
    getSnapshot: () => ({
      activeModal: 'jimaku',
      subtitlePreview: '字幕テキスト',
      secondarySubtitlePreview: 'secondary',
      isOverlayInteractive: true,
      isOverSubtitle: true,
      invisiblePositionEditMode: false,
      overlayLayer: 'visible',
    }),
    logError: (payload) => {
      payloads.push(payload);
    },
  });

  controller.handleError(new Error('renderer boom'), {
    source: 'callback',
    action: 'onSubtitle',
  });

  assert.equal(dismissed, 1);
  assert.equal(restored, 1);
  assert.equal(shown.length, 1);
  assert.match(shown[0], /recovered/i);
  assert.equal(payloads.length, 1);

  const payload = payloads[0] as {
    context: { action: string };
    error: { message: string; stack: string | null };
    snapshot: { activeModal: string | null; subtitlePreview: string };
  };
  assert.equal(payload.context.action, 'onSubtitle');
  assert.equal(payload.snapshot.activeModal, 'jimaku');
  assert.equal(payload.snapshot.subtitlePreview, '字幕テキスト');
  assert.equal(payload.error.message, 'renderer boom');
  assert.ok(
    typeof payload.error.stack === 'string' && payload.error.stack.includes('renderer boom'),
  );
});

test('handleError normalizes non-Error values', () => {
  const payloads: unknown[] = [];

  const controller = createRendererRecoveryController({
    dismissActiveUi: () => {},
    restoreOverlayInteraction: () => {},
    showToast: () => {},
    getSnapshot: () => ({
      activeModal: null,
      subtitlePreview: '',
      secondarySubtitlePreview: '',
      isOverlayInteractive: false,
      isOverSubtitle: false,
      invisiblePositionEditMode: false,
      overlayLayer: 'invisible',
    }),
    logError: (payload) => {
      payloads.push(payload);
    },
  });

  controller.handleError({ code: 500, reason: 'timeout' }, { source: 'callback', action: 'modal' });

  const payload = payloads[0] as { error: { message: string; stack: string | null } };
  assert.equal(payload.error.message, JSON.stringify({ code: 500, reason: 'timeout' }));
  assert.equal(payload.error.stack, null);
});

test('nested recovery errors are ignored while current recovery is active', () => {
  const payloads: unknown[] = [];
  let restored = 0;

  let controllerRef: ReturnType<typeof createRendererRecoveryController> | null = null;

  const controller = createRendererRecoveryController({
    dismissActiveUi: () => {
      controllerRef?.handleError(new Error('nested'), { source: 'callback', action: 'nested' });
    },
    restoreOverlayInteraction: () => {
      restored += 1;
    },
    showToast: () => {},
    getSnapshot: () => ({
      activeModal: 'runtime-options',
      subtitlePreview: '',
      secondarySubtitlePreview: '',
      isOverlayInteractive: true,
      isOverSubtitle: false,
      invisiblePositionEditMode: true,
      overlayLayer: 'visible',
    }),
    logError: (payload) => {
      payloads.push(payload);
    },
  });
  controllerRef = controller;

  controller.handleError(new Error('outer'), { source: 'callback', action: 'outer' });

  assert.equal(payloads.length, 1);
  assert.equal(restored, 1);
});
