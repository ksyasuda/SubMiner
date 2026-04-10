import assert from 'node:assert/strict';
import test from 'node:test';
import { initializeOverlayAnkiIntegration, initializeOverlayRuntime } from './overlay-runtime-init';

function withPlatform(platform: NodeJS.Platform, run: () => void): void {
  const originalPlatform = process.platform;
  Object.defineProperty(process, 'platform', {
    value: platform,
    configurable: true,
  });

  try {
    run();
  } finally {
    Object.defineProperty(process, 'platform', {
      value: originalPlatform,
      configurable: true,
    });
  }
}

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

test('initializeOverlayAnkiIntegration can initialize Anki transport after overlay runtime already exists', () => {
  let createdIntegrations = 0;
  let startedIntegrations = 0;
  let setIntegrationCalls = 0;

  initializeOverlayAnkiIntegration({
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
      keepNoteId: 11,
      deleteNoteId: 12,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });

  assert.equal(createdIntegrations, 1);
  assert.equal(startedIntegrations, 1);
  assert.equal(setIntegrationCalls, 1);
});

test('initializeOverlayAnkiIntegration returns false when integration already exists', () => {
  let createdIntegrations = 0;
  let startedIntegrations = 0;
  let setIntegrationCalls = 0;

  const result = initializeOverlayAnkiIntegration({
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
    getAnkiIntegration: () => ({}),
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
      keepNoteId: 11,
      deleteNoteId: 12,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });

  assert.equal(result, false);
  assert.equal(createdIntegrations, 0);
  assert.equal(startedIntegrations, 0);
  assert.equal(setIntegrationCalls, 0);
});

test('initializeOverlayAnkiIntegration returns false when ankiConnect is disabled', () => {
  let createdIntegrations = 0;
  let startedIntegrations = 0;
  let setIntegrationCalls = 0;

  const result = initializeOverlayAnkiIntegration({
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
      keepNoteId: 11,
      deleteNoteId: 12,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
  });

  assert.equal(result, false);
  assert.equal(createdIntegrations, 0);
  assert.equal(startedIntegrations, 0);
  assert.equal(setIntegrationCalls, 0);
});

test('initializeOverlayRuntime can skip starting Anki integration transport', () => {
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
      keepNoteId: 7,
      deleteNoteId: 8,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
    shouldStartAnkiIntegration: () => false,
  });

  assert.equal(createdIntegrations, 1);
  assert.equal(startedIntegrations, 0);
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

test('initializeOverlayRuntime refreshes the current subtitle when tracker finds the target window again', () => {
  let subtitleRefreshCalls = 0;
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
    updateVisibleOverlayVisibility: () => {},
    refreshCurrentSubtitle: () => {
      subtitleRefreshCalls += 1;
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

  tracker.onWindowFound?.({ x: 100, y: 200, width: 1280, height: 720 });

  assert.equal(subtitleRefreshCalls, 1);
});

test('initializeOverlayRuntime hides overlay windows when tracker loses the target window', () => {
  const calls: string[] = [];
  const tracker = {
    onGeometryChange: null as ((...args: unknown[]) => void) | null,
    onWindowFound: null as ((...args: unknown[]) => void) | null,
    onWindowLost: null as (() => void) | null,
    onWindowFocusChange: null as ((focused: boolean) => void) | null,
    isTargetWindowMinimized: () => true,
    start: () => {},
  };
  const overlayWindows = [
    {
      hide: () => calls.push('hide-visible'),
    },
    {
      hide: () => calls.push('hide-modal'),
    },
  ];

  initializeOverlayRuntime({
    backendOverride: null,
    createMainWindow: () => {},
    registerGlobalShortcuts: () => {},
    updateVisibleOverlayBounds: () => {},
    isVisibleOverlayVisible: () => true,
    updateVisibleOverlayVisibility: () => {},
    refreshCurrentSubtitle: () => {},
    getOverlayWindows: () => overlayWindows as never,
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
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

  tracker.onWindowLost?.();

  assert.deepEqual(calls, ['hide-visible', 'hide-modal', 'sync-shortcuts']);
});

test('initializeOverlayRuntime preserves visible overlay on Windows tracker loss when target is not minimized', () => {
  withPlatform('win32', () => {
    const calls: string[] = [];
    const tracker = {
      onGeometryChange: null as ((...args: unknown[]) => void) | null,
      onWindowFound: null as ((...args: unknown[]) => void) | null,
      onWindowLost: null as (() => void) | null,
      onWindowFocusChange: null as ((focused: boolean) => void) | null,
      isTargetWindowMinimized: () => false,
      start: () => {},
    };
    const overlayWindows = [
      {
        hide: () => calls.push('hide-visible'),
      },
    ];

    initializeOverlayRuntime({
      backendOverride: null,
      createMainWindow: () => {},
      registerGlobalShortcuts: () => {},
      updateVisibleOverlayBounds: () => {},
      isVisibleOverlayVisible: () => true,
      updateVisibleOverlayVisibility: () => {
        calls.push('update-visible');
      },
      refreshCurrentSubtitle: () => {},
      getOverlayWindows: () => overlayWindows as never,
      syncOverlayShortcuts: () => {
        calls.push('sync-shortcuts');
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

    calls.length = 0;
    tracker.onWindowLost?.();

    assert.deepEqual(calls, ['sync-shortcuts']);
  });
});

test('initializeOverlayRuntime restores overlay bounds and visibility when tracker finds the target window again', () => {
  const bounds: Array<{ x: number; y: number; width: number; height: number }> = [];
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
    updateVisibleOverlayBounds: (geometry) => {
      bounds.push(geometry);
    },
    isVisibleOverlayVisible: () => true,
    updateVisibleOverlayVisibility: () => {
      visibilityRefreshCalls += 1;
    },
    refreshCurrentSubtitle: () => {},
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

  const restoredGeometry = { x: 100, y: 200, width: 1280, height: 720 };
  tracker.onWindowFound?.(restoredGeometry);

  assert.deepEqual(bounds, [restoredGeometry]);
  assert.equal(visibilityRefreshCalls, 2);
});
