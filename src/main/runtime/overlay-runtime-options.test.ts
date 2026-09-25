import test from 'node:test';
import assert from 'node:assert/strict';
import { createBuildInitializeOverlayRuntimeOptionsHandler } from './overlay-runtime-options';

test('build initialize overlay runtime options maps dependencies', () => {
  const calls: string[] = [];
  const buildOptions = createBuildInitializeOverlayRuntimeOptionsHandler({
    getBackendOverride: () => 'x11',
    createMainWindow: () => calls.push('create-main'),
    registerGlobalShortcuts: () => calls.push('register-shortcuts'),
    updateVisibleOverlayBounds: () => calls.push('update-visible-bounds'),
    isVisibleOverlayVisible: () => true,
    updateVisibleOverlayVisibility: () => calls.push('update-visible'),
    refreshCurrentSubtitle: () => calls.push('refresh-subtitle'),
    getOverlayWindows: () => [],
    syncOverlayShortcuts: () => calls.push('sync-shortcuts'),
    setWindowTracker: () => calls.push('set-tracker'),
    getResolvedConfig: () => ({}),
    getSubtitleTimingTracker: () => null,
    getMpvClient: () => null,
    getMpvSocketPath: () => '/tmp/mpv.sock',
    getRuntimeOptionsManager: () => null,
    setAnkiIntegration: () => calls.push('set-anki'),
    showDesktopNotification: () => calls.push('notify'),
    showOverlayNotification: () => calls.push('show-overlay'),
    dismissOverlayNotification: () => calls.push('dismiss-overlay'),
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
    shouldStartAnkiIntegration: () => true,
  });

  const options = buildOptions();
  assert.equal(options.backendOverride, 'x11');
  assert.equal(options.isVisibleOverlayVisible(), true);
  assert.equal(options.getMpvSocketPath(), '/tmp/mpv.sock');
  assert.equal(options.getKnownWordCacheStatePath(), '/tmp/known-words-cache.json');
  assert.equal(options.shouldStartAnkiIntegration(), true);
  options.createMainWindow();
  options.registerGlobalShortcuts();
  options.updateVisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  options.updateVisibleOverlayVisibility();
  options.refreshCurrentSubtitle?.();
  options.syncOverlayShortcuts();
  options.setWindowTracker(null);
  options.setAnkiIntegration(null);
  options.showDesktopNotification('title', {});
  options.showOverlayNotification?.({ title: 'title' });
  options.dismissOverlayNotification?.('notification-id');

  assert.deepEqual(calls, [
    'create-main',
    'register-shortcuts',
    'update-visible-bounds',
    'update-visible',
    'refresh-subtitle',
    'sync-shortcuts',
    'set-tracker',
    'set-anki',
    'notify',
    'show-overlay',
    'dismiss-overlay',
  ]);
});
