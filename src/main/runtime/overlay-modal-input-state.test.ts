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
    setFocusable: (focusable: boolean) => {
      calls.push(`focusable:${focusable}`);
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
    'focusable:true',
    'ignore:false',
    'top:true:screen-saver:3',
    'focus',
    'web-focus',
  ]);
  assert.deepEqual(calls, ['shortcuts:true', 'visibility']);
});

test('overlay modal input state restores main window focus on deactivation', () => {
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
    restoreMainWindowFocus: () => {
      calls.push('restore-focus');
    },
  });

  state.handleModalInputStateChange(true);
  calls.length = 0;

  state.handleModalInputStateChange(false);
  assert.deepEqual(calls, ['shortcuts:false', 'visibility', 'restore-focus', 'visibility']);
});

test('overlay modal input state schedules visibility settle burst after focus restore', () => {
  const modalWindow = createModalWindow();
  const calls: string[] = [];
  const scheduled: Array<{ delayMs: number; callback: () => void }> = [];
  const state = createOverlayModalInputState({
    getModalWindow: () => modalWindow as never,
    syncOverlayShortcutsForModal: () => {},
    syncOverlayVisibilityForModal: () => {
      calls.push('visibility');
    },
    restoreMainWindowFocus: () => {
      calls.push('restore-focus');
    },
    schedulePostRestoreVisibilitySync: (callback, delayMs) => {
      scheduled.push({ callback, delayMs });
      return scheduled.length as never;
    },
    clearPostRestoreVisibilitySync: () => {},
  });

  state.handleModalInputStateChange(true);
  calls.length = 0;

  state.handleModalInputStateChange(false);

  assert.deepEqual(
    scheduled.map((entry) => entry.delayMs),
    [50, 150, 300, 600, 1000],
  );
  for (const entry of scheduled) {
    entry.callback();
  }
  assert.deepEqual(calls, [
    'visibility',
    'restore-focus',
    'visibility',
    'visibility',
    'visibility',
    'visibility',
    'visibility',
    'visibility',
  ]);
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
