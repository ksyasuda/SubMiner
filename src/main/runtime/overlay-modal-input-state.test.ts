import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayModalInputState } from './overlay-modal-input-state';

function createModalWindow() {
  const calls: string[] = [];
  let destroyed = false;
  let focused = false;
  let webContentsFocused = false;

  return {
    calls,
    setDestroyed(next: boolean) {
      destroyed = next;
    },
    setFocused(next: boolean) {
      focused = next;
    },
    setWebContentsFocused(next: boolean) {
      webContentsFocused = next;
    },
    isDestroyed: () => destroyed,
    setIgnoreMouseEvents: (ignore: boolean) => {
      calls.push(`ignore:${ignore}`);
    },
    setAlwaysOnTop: (flag: boolean, level?: string, relativeLevel?: number) => {
      calls.push(`top:${flag}:${level ?? ''}:${relativeLevel ?? ''}`);
    },
    focus: () => {
      focused = true;
      calls.push('focus');
    },
    isFocused: () => focused,
    webContents: {
      isFocused: () => webContentsFocused,
      focus: () => {
        webContentsFocused = true;
        calls.push('web-focus');
      },
    },
  };
}

test('overlay modal input state activates modal window interactivity and syncs dependents', () => {
  const modalWindow = createModalWindow();
  const calls: string[] = [];
  const state = createOverlayModalInputState({
    getModalWindow: () => modalWindow as never,
    syncOverlayShortcutsForModal: (isActive) => {
      calls.push(`shortcuts:${isActive}`);
    },
    syncOverlayVisibilityForModal: () => {
      calls.push('visibility');
    },
  });

  state.handleModalInputStateChange(true);

  assert.equal(state.getModalInputExclusive(), true);
  assert.deepEqual(modalWindow.calls, [
    'ignore:false',
    'top:true:screen-saver:1',
    'focus',
    'web-focus',
  ]);
  assert.deepEqual(calls, ['shortcuts:true', 'visibility']);
});

test('overlay modal input state is idempotent for unchanged state', () => {
  const calls: string[] = [];
  const state = createOverlayModalInputState({
    getModalWindow: () => null,
    syncOverlayShortcutsForModal: (isActive) => {
      calls.push(`shortcuts:${isActive}`);
    },
    syncOverlayVisibilityForModal: () => {
      calls.push('visibility');
    },
  });

  state.handleModalInputStateChange(false);
  state.handleModalInputStateChange(true);
  state.handleModalInputStateChange(true);

  assert.equal(state.getModalInputExclusive(), true);
  assert.deepEqual(calls, ['shortcuts:true', 'visibility']);
});
