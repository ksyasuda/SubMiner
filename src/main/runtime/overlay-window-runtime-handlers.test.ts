import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayWindowRuntimeHandlers } from './overlay-window-runtime-handlers';

test('overlay window runtime handlers compose create/main/modal handlers', () => {
  let mainWindow: { kind: string } | null = null;
  let modalWindow: { kind: string } | null = null;
  let debugEnabled = false;
  const calls: string[] = [];

  const runtime = createOverlayWindowRuntimeHandlers({
    createOverlayWindowDeps: {
      createOverlayWindowCore: (kind) => ({ kind }),
      isDev: true,
      ensureOverlayWindowLevel: () => calls.push('ensure-level'),
      onRuntimeOptionsChanged: () => calls.push('runtime-options-changed'),
      setOverlayDebugVisualizationEnabled: (enabled) => {
        debugEnabled = enabled;
      },
      isOverlayVisible: (kind) => kind === 'visible',
      tryHandleOverlayShortcutLocalFallback: () => false,
      onWindowClosed: (kind) => calls.push(`closed:${kind}`),
    },
    setMainWindow: (window) => {
      mainWindow = window;
    },
    setModalWindow: (window) => {
      modalWindow = window;
    },
  });

  assert.deepEqual(runtime.createOverlayWindow('visible'), { kind: 'visible' });
  assert.deepEqual(runtime.createOverlayWindow('modal'), { kind: 'modal' });

  assert.deepEqual(runtime.createMainWindow(), { kind: 'visible' });
  assert.deepEqual(mainWindow, { kind: 'visible' });

  assert.deepEqual(runtime.createModalWindow(), { kind: 'modal' });
  assert.deepEqual(modalWindow, { kind: 'modal' });

  assert.equal(debugEnabled, false);
  assert.deepEqual(calls, []);
});
