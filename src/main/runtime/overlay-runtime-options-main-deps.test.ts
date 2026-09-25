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
    },
    overlayVisibilityRuntime: {
      updateVisibleOverlayVisibility: () => calls.push('update-visible'),
    },
    refreshCurrentSubtitle: () => calls.push('refresh-subtitle'),
    overlayShortcutsRuntime: {
      syncOverlayShortcuts: () => calls.push('sync-shortcuts'),
    },
    createMainWindow: () => calls.push('create-main'),
    registerGlobalShortcuts: () => calls.push('register-shortcuts'),
    updateVisibleOverlayBounds: () => calls.push('visible-bounds'),
    getOverlayWindows: () => [],
    getResolvedConfig: () => ({}),
    showDesktopNotification: () => calls.push('notify'),
    showOverlayNotification: () => calls.push('show-overlay'),
    dismissOverlayNotification: () => calls.push('dismiss-overlay'),
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: true,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
    shouldStartAnkiIntegration: () => false,
  });

  const deps = build();
  assert.equal(deps.getBackendOverride(), 'x11');
  assert.equal(deps.isVisibleOverlayVisible(), true);
  assert.equal(deps.getMpvSocketPath(), '/tmp/mpv.sock');
  assert.equal(deps.getKnownWordCacheStatePath(), '/tmp/known-words-cache.json');
  assert.equal(deps.shouldStartAnkiIntegration(), false);

  deps.createMainWindow();
  deps.registerGlobalShortcuts();
  deps.updateVisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  deps.updateVisibleOverlayVisibility();
  deps.refreshCurrentSubtitle?.();
  deps.syncOverlayShortcuts();
  deps.showDesktopNotification('title', {});
  deps.showOverlayNotification?.({ title: 'title' });
  deps.dismissOverlayNotification?.('notification-id');

  const tracker = {
    close: () => {},
    getWindowGeometry: () => null,
  } as unknown as BaseWindowTracker;
  deps.setWindowTracker(tracker);
  deps.setAnkiIntegration({ id: 'anki' });

  assert.deepEqual(calls, [
    'create-main',
    'register-shortcuts',
    'visible-bounds',
    'update-visible',
    'refresh-subtitle',
    'sync-shortcuts',
    'notify',
    'show-overlay',
    'dismiss-overlay',
  ]);
  assert.equal(appState.windowTracker, tracker);
  assert.deepEqual(appState.ankiIntegration, { id: 'anki' });
});
