import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveOverlayLayerFromArgv } from './preload-args';

test('resolveOverlayLayerFromArgv returns null when argv is unavailable', () => {
  assert.equal(resolveOverlayLayerFromArgv(null), null);
});

test('resolveOverlayLayerFromArgv returns parsed overlay layer when present', () => {
  assert.equal(resolveOverlayLayerFromArgv(['electron', '--overlay-layer=modal']), 'modal');
  assert.equal(resolveOverlayLayerFromArgv(['electron', '--overlay-layer=visible']), 'visible');
});

test('resolveOverlayLayerFromArgv ignores unsupported overlay layers', () => {
  assert.equal(resolveOverlayLayerFromArgv(['electron', '--overlay-layer=secondary']), null);
});
