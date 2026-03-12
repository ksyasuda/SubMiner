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
      options.onRuntimeOptionsChanged();
      options.setOverlayDebugVisualizationEnabled(true);
      options.onWindowClosed(kind);
      return window;
    },
    isDev: true,
    ensureOverlayWindowLevel: () => {},
    onRuntimeOptionsChanged: () => calls.push('runtime-options'),
    setOverlayDebugVisualizationEnabled: (enabled) => calls.push(`debug:${enabled}`),
    isOverlayVisible: (kind) => kind === 'visible',
    tryHandleOverlayShortcutLocalFallback: () => false,
    forwardTabToMpv: () => calls.push('forward-tab'),
    onWindowClosed: (kind) => calls.push(`closed:${kind}`),
    getYomitanSession: () => yomitanSession,
  });

  assert.equal(createOverlayWindow('visible'), window);
  assert.deepEqual(calls, [
    'kind:visible',
    'forward-tab',
    'runtime-options',
    'debug:true',
    'closed:visible',
  ]);
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
