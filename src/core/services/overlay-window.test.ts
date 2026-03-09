import assert from 'node:assert/strict';
import test from 'node:test';
import {
  handleOverlayWindowBeforeInputEvent,
  isTabInputForMpvForwarding,
} from './overlay-window-input';

test('isTabInputForMpvForwarding matches bare Tab keydown only', () => {
  assert.equal(
    isTabInputForMpvForwarding({
      type: 'keyDown',
      key: 'Tab',
      code: 'Tab',
    } as Electron.Input),
    true,
  );
  assert.equal(
    isTabInputForMpvForwarding({
      type: 'keyDown',
      key: 'Tab',
      code: 'Tab',
      shift: true,
    } as Electron.Input),
    false,
  );
  assert.equal(
    isTabInputForMpvForwarding({
      type: 'keyUp',
      key: 'Tab',
      code: 'Tab',
    } as Electron.Input),
    false,
  );
});

test('handleOverlayWindowBeforeInputEvent forwards Tab to mpv for visible overlays', () => {
  const calls: string[] = [];

  const handled = handleOverlayWindowBeforeInputEvent({
    kind: 'visible',
    windowVisible: true,
    input: {
      type: 'keyDown',
      key: 'Tab',
      code: 'Tab',
    } as Electron.Input,
    preventDefault: () => calls.push('prevent-default'),
    sendKeyboardModeToggleRequested: () => calls.push('keyboard-mode'),
    sendLookupWindowToggleRequested: () => calls.push('lookup-toggle'),
    tryHandleOverlayShortcutLocalFallback: () => {
      calls.push('fallback');
      return false;
    },
    forwardTabToMpv: () => calls.push('forward-tab'),
  });

  assert.equal(handled, true);
  assert.deepEqual(calls, ['prevent-default', 'forward-tab']);
});

test('handleOverlayWindowBeforeInputEvent leaves modal Tab handling alone', () => {
  const calls: string[] = [];

  const handled = handleOverlayWindowBeforeInputEvent({
    kind: 'modal',
    windowVisible: true,
    input: {
      type: 'keyDown',
      key: 'Tab',
      code: 'Tab',
    } as Electron.Input,
    preventDefault: () => calls.push('prevent-default'),
    sendKeyboardModeToggleRequested: () => calls.push('keyboard-mode'),
    sendLookupWindowToggleRequested: () => calls.push('lookup-toggle'),
    tryHandleOverlayShortcutLocalFallback: () => {
      calls.push('fallback');
      return false;
    },
    forwardTabToMpv: () => calls.push('forward-tab'),
  });

  assert.equal(handled, false);
  assert.deepEqual(calls, []);
});
