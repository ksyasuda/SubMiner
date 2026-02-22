import assert from 'node:assert/strict';
import test from 'node:test';
import type { BaseWindowTracker } from '../../window-trackers';
import { createBuildInitializeOverlayRuntimeMainDepsHandler } from './overlay-runtime-options-main-deps';

test('overlay runtime main deps builder maps runtime state and callbacks', () => {
  const calls: string[] = [];
  const appState = {
    backendOverride: 'x11' as string | null,
    windowTracker: null as BaseWindowTracker | null,
    subtitleTimingTracker: { id: 'tracker' } as unknown,
    mpvClient: null as { send?: (payload: { command: string[] }) => void } | null,
    mpvSocketPath: '/tmp/mpv.sock',
    runtimeOptionsManager: null,
    ankiIntegration: null as unknown,
  };

  const build = createBuildInitializeOverlayRuntimeMainDepsHandler({
    appState,
    overlayManager: {
      getVisibleOverlayVisible: () => true,
      getInvisibleOverlayVisible: () => false,
    },
    overlayVisibilityRuntime: {
      updateVisibleOverlayVisibility: () => calls.push('update-visible'),
      updateInvisibleOverlayVisibility: () => calls.push('update-invisible'),
    },
    overlayShortcutsRuntime: {
      syncOverlayShortcuts: () => calls.push('sync-shortcuts'),
    },
    getInitialInvisibleOverlayVisibility: () => true,
    createMainWindow: () => calls.push('create-main'),
    createInvisibleWindow: () => calls.push('create-invisible'),
    registerGlobalShortcuts: () => calls.push('register-shortcuts'),
    updateVisibleOverlayBounds: () => calls.push('visible-bounds'),
    updateInvisibleOverlayBounds: () => calls.push('invisible-bounds'),
    getOverlayWindows: () => [],
    getResolvedConfig: () => ({}),
    showDesktopNotification: () => calls.push('notify'),
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: true,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });

  const deps = build();
  assert.equal(deps.getBackendOverride(), 'x11');
  assert.equal(deps.getInitialInvisibleOverlayVisibility(), true);
  assert.equal(deps.isVisibleOverlayVisible(), true);
  assert.equal(deps.isInvisibleOverlayVisible(), false);
  assert.equal(deps.getMpvSocketPath(), '/tmp/mpv.sock');
  assert.equal(deps.getKnownWordCacheStatePath(), '/tmp/known-words-cache.json');

  deps.createMainWindow();
  deps.createInvisibleWindow();
  deps.registerGlobalShortcuts();
  deps.updateVisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  deps.updateInvisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  deps.updateVisibleOverlayVisibility();
  deps.updateInvisibleOverlayVisibility();
  deps.syncOverlayShortcuts();
  deps.showDesktopNotification('title', {});

  const tracker = {
    close: () => {},
    getWindowGeometry: () => null,
  } as unknown as BaseWindowTracker;
  deps.setWindowTracker(tracker);
  deps.setAnkiIntegration({ id: 'anki' });

  assert.deepEqual(calls, [
    'create-main',
    'create-invisible',
    'register-shortcuts',
    'visible-bounds',
    'invisible-bounds',
    'update-visible',
    'update-invisible',
    'sync-shortcuts',
    'notify',
  ]);
  assert.equal(appState.windowTracker, tracker);
  assert.deepEqual(appState.ankiIntegration, { id: 'anki' });
});
