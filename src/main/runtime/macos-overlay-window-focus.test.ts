import assert from 'node:assert/strict';
import test from 'node:test';
import { focusMacOSOverlayWindow } from './macos-overlay-window-focus';

function createOverlayWindowStub(
  overrides: Partial<{
    destroyed: boolean;
    visible: boolean;
    focused: boolean;
    webContentsFocused: boolean;
  }> = {},
  calls: string[] = [],
) {
  return {
    isDestroyed: () => overrides.destroyed ?? false,
    isVisible: () => overrides.visible ?? true,
    isFocused: () => overrides.focused ?? false,
    setFocusable: (focusable: boolean) => {
      calls.push(`setFocusable:${focusable}`);
    },
    focus: () => {
      calls.push('window.focus');
    },
    webContents: {
      isFocused: () => overrides.webContentsFocused ?? false,
      focus: () => {
        calls.push('webContents.focus');
      },
    },
  };
}

test('focusMacOSOverlayWindow activates the overlay window on macOS', () => {
  const calls: string[] = [];
  const overlayWindow = createOverlayWindowStub({}, calls);

  focusMacOSOverlayWindow({
    platform: 'darwin',
    getOverlayWindow: () => overlayWindow,
    stealAppFocus: () => calls.unshift('app.focus'),
    warn: () => {},
  });

  assert.deepEqual(calls, ['app.focus', 'setFocusable:true', 'window.focus', 'webContents.focus']);
});

test('focusMacOSOverlayWindow skips re-focusing the web contents when already focused', () => {
  const calls: string[] = [];
  const overlayWindow = createOverlayWindowStub({ webContentsFocused: true }, calls);

  focusMacOSOverlayWindow({
    platform: 'darwin',
    getOverlayWindow: () => overlayWindow,
    stealAppFocus: () => calls.unshift('app.focus'),
    warn: () => {},
  });

  assert.deepEqual(calls, ['app.focus', 'setFocusable:true', 'window.focus']);
});

test('focusMacOSOverlayWindow no-ops on non-macOS platforms', () => {
  const calls: string[] = [];

  focusMacOSOverlayWindow({
    platform: 'win32',
    getOverlayWindow: () => createOverlayWindowStub({}, calls),
    stealAppFocus: () => calls.push('app.focus'),
    warn: () => {},
  });

  assert.deepEqual(calls, []);
});

test('focusMacOSOverlayWindow no-ops when the overlay is already focused', () => {
  const calls: string[] = [];

  focusMacOSOverlayWindow({
    platform: 'darwin',
    getOverlayWindow: () => createOverlayWindowStub({ focused: true }, calls),
    stealAppFocus: () => calls.push('app.focus'),
    warn: () => {},
  });

  assert.deepEqual(calls, []);
});

test('focusMacOSOverlayWindow no-ops when the overlay is hidden, destroyed, or missing', () => {
  const calls: string[] = [];

  focusMacOSOverlayWindow({
    platform: 'darwin',
    getOverlayWindow: () => createOverlayWindowStub({ visible: false }, calls),
    stealAppFocus: () => calls.push('app.focus'),
    warn: () => {},
  });
  focusMacOSOverlayWindow({
    platform: 'darwin',
    getOverlayWindow: () => createOverlayWindowStub({ destroyed: true }, calls),
    stealAppFocus: () => calls.push('app.focus'),
    warn: () => {},
  });
  focusMacOSOverlayWindow({
    platform: 'darwin',
    getOverlayWindow: () => null,
    stealAppFocus: () => calls.push('app.focus'),
    warn: () => {},
  });

  assert.deepEqual(calls, []);
});

test('focusMacOSOverlayWindow still focuses the window when stealing app focus throws', () => {
  const calls: string[] = [];
  const overlayWindow = createOverlayWindowStub({}, calls);

  focusMacOSOverlayWindow({
    platform: 'darwin',
    getOverlayWindow: () => overlayWindow,
    stealAppFocus: () => {
      throw new Error('steal failed');
    },
    warn: (message) => calls.push(`warn:${message}`),
  });

  assert.deepEqual(calls, [
    'warn:Failed to steal app focus for overlay window',
    'setFocusable:true',
    'window.focus',
    'webContents.focus',
  ]);
});
