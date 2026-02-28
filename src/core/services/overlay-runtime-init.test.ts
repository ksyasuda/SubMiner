import assert from 'node:assert/strict';
import test from 'node:test';
import { initializeOverlayRuntime } from './overlay-runtime-init';

test('initializeOverlayRuntime skips Anki integration when ankiConnect.enabled is false', () => {
  let createdIntegrations = 0;
  let startedIntegrations = 0;
  let setIntegrationCalls = 0;

  initializeOverlayRuntime({
    backendOverride: null,
    createMainWindow: () => {},
    registerGlobalShortcuts: () => {},
    updateVisibleOverlayBounds: () => {},
    isVisibleOverlayVisible: () => false,
    updateVisibleOverlayVisibility: () => {},
    getOverlayWindows: () => [],
    syncOverlayShortcuts: () => {},
    setWindowTracker: () => {},
    getMpvSocketPath: () => '/tmp/mpv.sock',
    createWindowTracker: () => null,
    getResolvedConfig: () => ({
      ankiConnect: { enabled: false } as never,
    }),
    getSubtitleTimingTracker: () => ({}),
    getMpvClient: () => ({
      send: () => {},
    }),
    getRuntimeOptionsManager: () => ({
      getEffectiveAnkiConnectConfig: (config) => config as never,
    }),
    createAnkiIntegration: () => {
      createdIntegrations += 1;
      return {
        start: () => {
          startedIntegrations += 1;
        },
      };
    },
    setAnkiIntegration: () => {
      setIntegrationCalls += 1;
    },
    showDesktopNotification: () => {},
    createFieldGroupingCallback: () =>
      async () => ({
        keepNoteId: 1,
        deleteNoteId: 2,
        deleteDuplicate: false,
        cancelled: false,
      }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });

  assert.equal(createdIntegrations, 0);
  assert.equal(startedIntegrations, 0);
  assert.equal(setIntegrationCalls, 0);
});

test('initializeOverlayRuntime starts Anki integration when ankiConnect.enabled is true', () => {
  let createdIntegrations = 0;
  let startedIntegrations = 0;
  let setIntegrationCalls = 0;

  initializeOverlayRuntime({
    backendOverride: null,
    createMainWindow: () => {},
    registerGlobalShortcuts: () => {},
    updateVisibleOverlayBounds: () => {},
    isVisibleOverlayVisible: () => false,
    updateVisibleOverlayVisibility: () => {},
    getOverlayWindows: () => [],
    syncOverlayShortcuts: () => {},
    setWindowTracker: () => {},
    getMpvSocketPath: () => '/tmp/mpv.sock',
    createWindowTracker: () => null,
    getResolvedConfig: () => ({
      ankiConnect: { enabled: true } as never,
    }),
    getSubtitleTimingTracker: () => ({}),
    getMpvClient: () => ({
      send: () => {},
    }),
    getRuntimeOptionsManager: () => ({
      getEffectiveAnkiConnectConfig: (config) => config as never,
    }),
    createAnkiIntegration: (args) => {
      createdIntegrations += 1;
      assert.equal(args.config.enabled, true);
      return {
        start: () => {
          startedIntegrations += 1;
        },
      };
    },
    setAnkiIntegration: () => {
      setIntegrationCalls += 1;
    },
    showDesktopNotification: () => {},
    createFieldGroupingCallback: () =>
      async () => ({
        keepNoteId: 3,
        deleteNoteId: 4,
        deleteDuplicate: false,
        cancelled: false,
      }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });

  assert.equal(createdIntegrations, 1);
  assert.equal(startedIntegrations, 1);
  assert.equal(setIntegrationCalls, 1);
});
