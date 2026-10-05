import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createCreateMainWindowHandler,
  createCreateOverlayWindowHandler,
} from './overlay-window-factory';

test('create overlay window handler applies the linux X11 fullscreen flag only to the visible window', () => {
  const seen: Array<{ kind: string; linuxX11FullscreenOverlay?: boolean }> = [];
  const createOverlayWindow = createCreateOverlayWindowHandler({
    createOverlayWindowCore: (kind, options) => {
      seen.push({ kind, linuxX11FullscreenOverlay: options.linuxX11FullscreenOverlay });
      return {};
    },
    isDev: false,
    ensureOverlayWindowLevel: () => {},
    onRuntimeOptionsChanged: () => {},
    setOverlayDebugVisualizationEnabled: () => {},
    isOverlayVisible: () => false,
    tryHandleOverlayShortcutLocalFallback: () => false,
    forwardTabToMpv: () => {},
    onWindowClosed: () => {},
    getLinuxX11FullscreenOverlay: () => true,
  });

  createOverlayWindow('visible');
  createOverlayWindow('modal');

  assert.deepEqual(seen, [
    { kind: 'visible', linuxX11FullscreenOverlay: true },
    { kind: 'modal', linuxX11FullscreenOverlay: undefined },
  ]);
});

test('create overlay window handler defaults the yomitan session to null', () => {
  let session: unknown = 'unset';
  const createOverlayWindow = createCreateOverlayWindowHandler({
    createOverlayWindowCore: (_kind, options) => {
      session = options.yomitanSession;
      return {};
    },
    isDev: false,
    ensureOverlayWindowLevel: () => {},
    onRuntimeOptionsChanged: () => {},
    setOverlayDebugVisualizationEnabled: () => {},
    isOverlayVisible: () => false,
    tryHandleOverlayShortcutLocalFallback: () => false,
    forwardTabToMpv: () => {},
    onWindowClosed: () => {},
  });

  createOverlayWindow('visible');

  assert.equal(session, null);
});

test('create main window handler stores visible window', () => {
  const calls: string[] = [];
  const visibleWindow = { id: 'visible' };
  let mainWindow: typeof visibleWindow | null = null;
  const createMainWindow = createCreateMainWindowHandler({
    getMainWindow: () => mainWindow,
    isWindowDestroyed: () => false,
    createOverlayWindow: (kind) => {
      calls.push(`create:${kind}`);
      return visibleWindow;
    },
    setMainWindow: (window) => {
      mainWindow = window;
      calls.push(`set:${(window as { id: string }).id}`);
    },
  });

  assert.equal(createMainWindow(), visibleWindow);
  assert.deepEqual(calls, ['create:visible', 'set:visible']);
});

test('create main window handler reuses an existing live visible window', () => {
  const calls: string[] = [];
  const existingWindow = { id: 'existing' };
  const createMainWindow = createCreateMainWindowHandler({
    getMainWindow: () => existingWindow,
    isWindowDestroyed: () => false,
    createOverlayWindow: (kind) => {
      calls.push(`create:${kind}`);
      return { id: 'created' };
    },
    setMainWindow: (window) => calls.push(`set:${(window as { id: string }).id}`),
  });

  assert.equal(createMainWindow(), existingWindow);
  assert.deepEqual(calls, []);
});
