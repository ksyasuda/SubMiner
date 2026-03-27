import assert from 'node:assert/strict';
import test from 'node:test';
import { composeOverlayWindowHandlers } from './overlay-window-composer';

test('composeOverlayWindowHandlers returns overlay window handlers', () => {
  let mainWindow: { kind: string } | null = null;
  let modalWindow: { kind: string } | null = null;

  const handlers = composeOverlayWindowHandlers<{ kind: string }>({
    createOverlayWindowDeps: {
      createOverlayWindowCore: (kind) => ({ kind }),
      isDev: false,
      ensureOverlayWindowLevel: () => {},
      onRuntimeOptionsChanged: () => {},
      setOverlayDebugVisualizationEnabled: () => {},
      isOverlayVisible: (kind) => kind === 'visible',
      tryHandleOverlayShortcutLocalFallback: () => false,
      forwardTabToMpv: () => {},
      onWindowClosed: () => {},
      getYomitanSession: () => null,
    },
    setMainWindow: (window) => {
      mainWindow = window;
    },
    setModalWindow: (window) => {
      modalWindow = window;
    },
  });

  assert.deepEqual(handlers.createMainWindow(), { kind: 'visible' });
  assert.deepEqual(mainWindow, { kind: 'visible' });
  assert.deepEqual(handlers.createModalWindow(), { kind: 'modal' });
  assert.deepEqual(modalWindow, { kind: 'modal' });
});
