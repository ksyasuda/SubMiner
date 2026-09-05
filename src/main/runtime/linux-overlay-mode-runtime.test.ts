import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { test } from 'node:test';
import { createLinuxOverlayModeRuntime } from './linux-overlay-mode-runtime';

class TestWindow extends EventEmitter {
  destroyed = false;
  hidden = false;
  isDestroyed() {
    return this.destroyed;
  }
  hide() {
    this.hidden = true;
  }
  destroy() {
    this.destroyed = true;
  }
  finishClose() {
    this.emit('closed');
  }
}

function fixture() {
  const initial = new TestWindow();
  const state: { window: TestWindow | null; visible: boolean; creates: number; refreshes: number } =
    {
      window: initial,
      visible: true,
      creates: 0,
      refreshes: 0,
    };
  const runtime = createLinuxOverlayModeRuntime({
    isEnabled: () => true,
    isVisible: () => state.visible,
    getWindow: () => state.window,
    clearWindow: () => {
      state.window = null;
    },
    createWindow: () => {
      state.creates += 1;
      state.window = new TestWindow();
    },
    refreshWindow: () => {
      state.refreshes += 1;
    },
    now: () => 42,
    logDebug: () => {},
  });
  return { initial, state, runtime };
}

test('Linux mode transition waits for close before replacing and refreshing the window', () => {
  const { initial, state, runtime } = fixture();
  runtime.ownerBindingKey = 'old-owner';
  runtime.sync(true);
  assert.equal(runtime.mode, 'fullscreen-override');
  assert.equal(runtime.fullscreenChangedAtMs, 42);
  assert.equal(runtime.ownerBindingKey, null);
  assert.equal(initial.hidden, true);
  assert.equal(state.creates, 0);
  initial.finishClose();
  assert.equal(state.creates, 1);
  assert.equal(state.refreshes, 1);
  runtime.sync(true);
  assert.equal(state.creates, 1);
});

test('an older close callback cannot clear or replace a newer overlay', () => {
  const { initial, state, runtime } = fixture();
  runtime.sync(true);
  runtime.sync(false);
  const replacement = state.window;
  assert.equal(state.creates, 1);
  initial.finishClose();
  assert.equal(state.window, replacement);
  assert.equal(state.creates, 1);
  assert.equal(runtime.mode, 'managed');
});

test('hiding or cancelling a transition prevents delayed window creation', () => {
  for (const cancel of [false, true]) {
    const { initial, state, runtime } = fixture();
    runtime.sync(true);
    if (cancel) runtime.cancelPendingTransition();
    else state.visible = false;
    initial.finishClose();
    assert.equal(state.creates, 0);
    assert.equal(state.window, null);
    state.visible = true;
    runtime.sync(true);
    assert.equal(state.creates, 1);
  }
});
