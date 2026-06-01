import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createCreateMainWindowHandler,
  createCreateModalWindowHandler,
  createCreateOverlayWindowHandler,
} from './overlay-window-factory';

test('create overlay window handler forwards options and kind', () => {
  const calls: string[] = [];
  const window = { id: 1 };
  const yomitanSession = { id: 'session' } as never;
  const createOverlayWindow = createCreateOverlayWindowHandler({
    createOverlayWindowCore: (kind, options) => {
      calls.push(`kind:${kind}`);
      assert.equal(options.isDev, true);
      assert.equal(options.isOverlayVisible('visible'), true);
      assert.equal(options.isOverlayVisible('modal'), false);
      assert.equal(options.yomitanSession, yomitanSession);
      options.forwardTabToMpv();
      options.onVisibleWindowFocused?.();
      options.onRuntimeOptionsChanged();
      options.setOverlayDebugVisualizationEnabled(true);
      options.onWindowClosed(kind, window);
      return window;
    },
    isDev: true,
    ensureOverlayWindowLevel: () => {},
    onRuntimeOptionsChanged: () => calls.push('runtime-options'),
    setOverlayDebugVisualizationEnabled: (enabled) => calls.push(`debug:${enabled}`),
    isOverlayVisible: (kind) => kind === 'visible',
    tryHandleOverlayShortcutLocalFallback: () => false,
    forwardTabToMpv: () => calls.push('forward-tab'),
    onVisibleWindowFocused: () => calls.push('visible-focus'),
    onWindowClosed: (kind, closedWindow) =>
      calls.push(`closed:${kind}:${(closedWindow as { id: number }).id}`),
    getYomitanSession: () => yomitanSession,
  });

  assert.equal(createOverlayWindow('visible'), window);
  assert.deepEqual(calls, [
    'kind:visible',
    'forward-tab',
    'visible-focus',
    'runtime-options',
    'debug:true',
    'closed:visible:1',
  ]);
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

test('create modal window handler stores modal window', () => {
  const calls: string[] = [];
  const modalWindow = { id: 'modal' };
  const createModalWindow = createCreateModalWindowHandler({
    createOverlayWindow: (kind) => {
      calls.push(`create:${kind}`);
      return modalWindow;
    },
    setModalWindow: (window) => calls.push(`set:${(window as { id: string }).id}`),
  });

  assert.equal(createModalWindow(), modalWindow);
  assert.deepEqual(calls, ['create:modal', 'set:modal']);
});
