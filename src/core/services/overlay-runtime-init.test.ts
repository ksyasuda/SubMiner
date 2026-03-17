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
    createFieldGroupingCallback: () => async () => ({
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
    createFieldGroupingCallback: () => async () => ({
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

test('initializeOverlayRuntime merges shared ai config with Anki overrides', () => {
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
      ankiConnect: {
        enabled: true,
        ai: {
          enabled: true,
          model: 'openrouter/anki-model',
          systemPrompt: 'Translate mined sentence text.',
        },
      } as never,
      ai: {
        enabled: true,
        apiKey: 'shared-key',
        baseUrl: 'https://openrouter.ai/api',
        model: 'openrouter/shared-model',
        systemPrompt: 'Legacy shared prompt.',
        requestTimeoutMs: 15000,
      },
    }),
    getSubtitleTimingTracker: () => ({}),
    getMpvClient: () => ({
      send: () => {},
    }),
    getRuntimeOptionsManager: () => ({
      getEffectiveAnkiConnectConfig: (config) => config as never,
    }),
    createAnkiIntegration: (args) => {
      assert.equal(args.aiConfig.apiKey, 'shared-key');
      assert.equal(args.aiConfig.baseUrl, 'https://openrouter.ai/api');
      assert.equal(args.aiConfig.model, 'openrouter/anki-model');
      assert.equal(args.aiConfig.systemPrompt, 'Translate mined sentence text.');
      return {
        start: () => {},
      };
    },
    setAnkiIntegration: () => {},
    showDesktopNotification: () => {},
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 5,
      deleteNoteId: 6,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });
});

test('initializeOverlayRuntime re-syncs overlay shortcuts when tracker focus changes', () => {
  let syncCalls = 0;
  const tracker = {
    onGeometryChange: null as ((...args: unknown[]) => void) | null,
    onWindowFound: null as ((...args: unknown[]) => void) | null,
    onWindowLost: null as (() => void) | null,
    onWindowFocusChange: null as ((focused: boolean) => void) | null,
    start: () => {},
  };

  initializeOverlayRuntime({
    backendOverride: null,
    createMainWindow: () => {},
    registerGlobalShortcuts: () => {},
    updateVisibleOverlayBounds: () => {},
    isVisibleOverlayVisible: () => false,
    updateVisibleOverlayVisibility: () => {},
    getOverlayWindows: () => [],
    syncOverlayShortcuts: () => {
      syncCalls += 1;
    },
    setWindowTracker: () => {},
    getMpvSocketPath: () => '/tmp/mpv.sock',
    createWindowTracker: () => tracker as never,
    getResolvedConfig: () => ({
      ankiConnect: { enabled: false } as never,
    }),
    getSubtitleTimingTracker: () => null,
    getMpvClient: () => null,
    getRuntimeOptionsManager: () => null,
    setAnkiIntegration: () => {},
    showDesktopNotification: () => {},
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });

  assert.equal(typeof tracker.onWindowFocusChange, 'function');
  tracker.onWindowFocusChange?.(true);
  assert.equal(syncCalls, 1);
});

test('initializeOverlayRuntime refreshes visible overlay when tracker focus changes while overlay is shown', () => {
  let visibilityRefreshCalls = 0;
  const tracker = {
    onGeometryChange: null as ((...args: unknown[]) => void) | null,
    onWindowFound: null as ((...args: unknown[]) => void) | null,
    onWindowLost: null as (() => void) | null,
    onWindowFocusChange: null as ((focused: boolean) => void) | null,
    start: () => {},
  };

  initializeOverlayRuntime({
    backendOverride: null,
    createMainWindow: () => {},
    registerGlobalShortcuts: () => {},
    updateVisibleOverlayBounds: () => {},
    isVisibleOverlayVisible: () => true,
    updateVisibleOverlayVisibility: () => {
      visibilityRefreshCalls += 1;
    },
    getOverlayWindows: () => [],
    syncOverlayShortcuts: () => {},
    setWindowTracker: () => {},
    getMpvSocketPath: () => '/tmp/mpv.sock',
    createWindowTracker: () => tracker as never,
    getResolvedConfig: () => ({
      ankiConnect: { enabled: false } as never,
    }),
    getSubtitleTimingTracker: () => null,
    getMpvClient: () => null,
    getRuntimeOptionsManager: () => null,
    setAnkiIntegration: () => {},
    showDesktopNotification: () => {},
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });

  tracker.onWindowFocusChange?.(true);

  assert.equal(visibilityRefreshCalls, 2);
});
