import test from 'node:test';
import assert from 'node:assert/strict';
import { createBuildInitializeOverlayRuntimeOptionsHandler } from './overlay-runtime-options';

test('build initialize overlay runtime options maps dependencies', () => {
  const calls: string[] = [];
  const buildOptions = createBuildInitializeOverlayRuntimeOptionsHandler({
    getBackendOverride: () => 'x11',
    getInitialInvisibleOverlayVisibility: () => true,
    createMainWindow: () => calls.push('create-main'),
    createInvisibleWindow: () => calls.push('create-invisible'),
    registerGlobalShortcuts: () => calls.push('register-shortcuts'),
    updateVisibleOverlayBounds: () => calls.push('update-visible-bounds'),
    updateInvisibleOverlayBounds: () => calls.push('update-invisible-bounds'),
    isVisibleOverlayVisible: () => true,
    isInvisibleOverlayVisible: () => false,
    updateVisibleOverlayVisibility: () => calls.push('update-visible'),
    updateInvisibleOverlayVisibility: () => calls.push('update-invisible'),
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
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });

  const options = buildOptions();
  assert.equal(options.backendOverride, 'x11');
  assert.equal(options.getInitialInvisibleOverlayVisibility(), true);
  assert.equal(options.isVisibleOverlayVisible(), true);
  assert.equal(options.isInvisibleOverlayVisible(), false);
  assert.equal(options.getMpvSocketPath(), '/tmp/mpv.sock');
  assert.equal(options.getKnownWordCacheStatePath(), '/tmp/known-words-cache.json');
  options.createMainWindow();
  options.createInvisibleWindow();
  options.registerGlobalShortcuts();
  options.updateVisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  options.updateInvisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  options.updateVisibleOverlayVisibility();
  options.updateInvisibleOverlayVisibility();
  options.syncOverlayShortcuts();
  options.setWindowTracker(null);
  options.setAnkiIntegration(null);
  options.showDesktopNotification('title', {});

  assert.deepEqual(calls, [
    'create-main',
    'create-invisible',
    'register-shortcuts',
    'update-visible-bounds',
    'update-invisible-bounds',
    'update-visible',
    'update-invisible',
    'sync-shortcuts',
    'set-tracker',
    'set-anki',
    'notify',
  ]);
});
