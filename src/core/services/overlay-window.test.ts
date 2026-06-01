import assert from 'node:assert/strict';
import test from 'node:test';
import { ensureOverlayWindowLevel } from './overlay-window';
import {
  handleOverlayWindowBeforeInputEvent,
  handleOverlayWindowBlurred,
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

test('handleOverlayWindowBlurred skips visible overlay restacking after manual hide', () => {
  const calls: string[] = [];

  const handled = handleOverlayWindowBlurred({
    kind: 'visible',
    windowVisible: true,
    isOverlayVisible: () => false,
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    moveWindowTop: () => {
      calls.push('move-top');
    },
  });

  assert.equal(handled, false);
  assert.deepEqual(calls, []);
});

test('handleOverlayWindowBlurred skips Windows visible overlay restacking after focus loss', () => {
  const calls: string[] = [];

  const handled = handleOverlayWindowBlurred({
    kind: 'visible',
    windowVisible: true,
    isOverlayVisible: () => true,
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    moveWindowTop: () => {
      calls.push('move-top');
    },
    platform: 'win32',
  });

  assert.equal(handled, false);
  assert.deepEqual(calls, []);
});

test('handleOverlayWindowBlurred notifies Windows visible overlay blur callback without restacking', () => {
  const calls: string[] = [];

  const handled = handleOverlayWindowBlurred({
    kind: 'visible',
    windowVisible: true,
    isOverlayVisible: () => true,
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    moveWindowTop: () => {
      calls.push('move-top');
    },
    onVisibleOverlayBlur: () => {
      calls.push('visible-blur');
    },
    platform: 'win32',
  });

  assert.equal(handled, false);
  assert.deepEqual(calls, ['visible-blur']);
});

test('handleOverlayWindowBlurred skips macOS visible overlay restacking after focus loss', () => {
  const calls: string[] = [];

  const handled = handleOverlayWindowBlurred({
    kind: 'visible',
    windowVisible: true,
    isOverlayVisible: () => true,
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    moveWindowTop: () => {
      calls.push('move-top');
    },
    platform: 'darwin',
  });

  assert.equal(handled, false);
  assert.deepEqual(calls, []);
});

test('handleOverlayWindowBlurred skips Linux visible overlay restacking after focus loss', () => {
  const calls: string[] = [];

  const handled = handleOverlayWindowBlurred({
    kind: 'visible',
    windowVisible: true,
    isOverlayVisible: () => true,
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    moveWindowTop: () => {
      calls.push('move-top');
    },
    platform: 'linux',
  });

  assert.equal(handled, false);
  assert.deepEqual(calls, []);
});

test('handleOverlayWindowBlurred notifies Linux visible overlay blur callback without restacking', () => {
  const calls: string[] = [];

  const handled = handleOverlayWindowBlurred({
    kind: 'visible',
    windowVisible: true,
    isOverlayVisible: () => true,
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    moveWindowTop: () => {
      calls.push('move-top');
    },
    onVisibleOverlayBlur: () => {
      calls.push('visible-blur');
    },
    platform: 'linux',
  });

  assert.equal(handled, false);
  assert.deepEqual(calls, ['visible-blur']);
});

test('handleOverlayWindowBlurred notifies macOS visible overlay blur callback without restacking', () => {
  const calls: string[] = [];

  const handled = handleOverlayWindowBlurred({
    kind: 'visible',
    windowVisible: true,
    isOverlayVisible: () => true,
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    moveWindowTop: () => {
      calls.push('move-top');
    },
    onVisibleOverlayBlur: () => {
      calls.push('visible-blur');
    },
    platform: 'darwin',
  });

  assert.equal(handled, false);
  assert.deepEqual(calls, ['visible-blur']);
});

test('handleOverlayWindowBlurred preserves modal window stacking', () => {
  const calls: string[] = [];

  assert.equal(
    handleOverlayWindowBlurred({
      kind: 'modal',
      windowVisible: true,
      isOverlayVisible: () => false,
      ensureOverlayWindowLevel: () => {
        calls.push('ensure-modal');
      },
      moveWindowTop: () => {
        calls.push('move-modal');
      },
    }),
    true,
  );

  assert.deepEqual(calls, ['ensure-modal']);
});

test('ensureOverlayWindowLevel promotes Linux overlay above fullscreen mpv without changing workspaces', () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'linux',
  });

  const calls: string[] = [];

  try {
    ensureOverlayWindowLevel({
      getTitle: () => 'SubMiner Overlay',
      moveTop: () => calls.push('move-top'),
      setAlwaysOnTop: (flag: boolean, level?: string, relativeLevel?: number) => {
        calls.push(`always-on-top:${flag}:${level ?? 'none'}:${relativeLevel ?? 0}`);
      },
      setVisibleOnAllWorkspaces: (flag: boolean, options?: { visibleOnFullScreen?: boolean }) => {
        calls.push(
          `all-workspaces:${flag}:${options?.visibleOnFullScreen === true ? 'fullscreen' : 'plain'}`,
        );
      },
    } as never);
  } finally {
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }

  assert.deepEqual(calls, [
    'always-on-top:true:screen-saver:1',
    'all-workspaces:true:fullscreen',
    'move-top',
  ]);
});
