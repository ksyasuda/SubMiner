import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createCreateInvisibleWindowHandler,
  createCreateMainWindowHandler,
  createCreateModalWindowHandler,
  createCreateOverlayWindowHandler,
  createCreateSecondaryWindowHandler,
} from './overlay-window-factory';

test('create overlay window handler forwards options and kind', () => {
  const calls: string[] = [];
  const window = { id: 1 };
  const createOverlayWindow = createCreateOverlayWindowHandler({
    createOverlayWindowCore: (kind, options) => {
      calls.push(`kind:${kind}`);
      assert.equal(options.isDev, true);
      assert.equal(options.overlayDebugVisualizationEnabled, false);
      assert.equal(options.isOverlayVisible('visible'), true);
      assert.equal(options.isOverlayVisible('invisible'), false);
      options.onRuntimeOptionsChanged();
      options.setOverlayDebugVisualizationEnabled(true);
      options.onWindowClosed(kind);
      return window;
    },
    isDev: true,
    getOverlayDebugVisualizationEnabled: () => false,
    ensureOverlayWindowLevel: () => {},
    onRuntimeOptionsChanged: () => calls.push('runtime-options'),
    setOverlayDebugVisualizationEnabled: (enabled) => calls.push(`debug:${enabled}`),
    isOverlayVisible: (kind) => kind === 'visible',
    tryHandleOverlayShortcutLocalFallback: () => false,
    onWindowClosed: (kind) => calls.push(`closed:${kind}`),
  });

  assert.equal(createOverlayWindow('visible'), window);
  assert.deepEqual(calls, ['kind:visible', 'runtime-options', 'debug:true', 'closed:visible']);
});

test('create main window handler stores visible window', () => {
  const calls: string[] = [];
  const visibleWindow = { id: 'visible' };
  const createMainWindow = createCreateMainWindowHandler({
    createOverlayWindow: (kind) => {
      calls.push(`create:${kind}`);
      return visibleWindow;
    },
    setMainWindow: (window) => calls.push(`set:${(window as { id: string }).id}`),
  });

  assert.equal(createMainWindow(), visibleWindow);
  assert.deepEqual(calls, ['create:visible', 'set:visible']);
});

test('create invisible window handler stores invisible window', () => {
  const calls: string[] = [];
  const invisibleWindow = { id: 'invisible' };
  const createInvisibleWindow = createCreateInvisibleWindowHandler({
    createOverlayWindow: (kind) => {
      calls.push(`create:${kind}`);
      return invisibleWindow;
    },
    setInvisibleWindow: (window) => calls.push(`set:${(window as { id: string }).id}`),
  });

  assert.equal(createInvisibleWindow(), invisibleWindow);
  assert.deepEqual(calls, ['create:invisible', 'set:invisible']);
});

test('create secondary window handler stores secondary window', () => {
  const calls: string[] = [];
  const secondaryWindow = { id: 'secondary' };
  const createSecondaryWindow = createCreateSecondaryWindowHandler({
    createOverlayWindow: (kind) => {
      calls.push(`create:${kind}`);
      return secondaryWindow;
    },
    setSecondaryWindow: (window) => calls.push(`set:${(window as { id: string }).id}`),
  });

  assert.equal(createSecondaryWindow(), secondaryWindow);
  assert.deepEqual(calls, ['create:secondary', 'set:secondary']);
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
