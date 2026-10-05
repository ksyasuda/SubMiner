import assert from 'node:assert/strict';
import test from 'node:test';
import type { WindowGeometry } from '../../types';
import {
  initializeOverlayAnkiIntegration,
  initializeOverlayRuntime,
  startOverlayWindowTracker,
} from './overlay-runtime-init';

type AnkiOptions = Parameters<typeof initializeOverlayAnkiIntegration>[0];
type RuntimeOptions = Parameters<typeof initializeOverlayRuntime>[0];
type CreateAnkiIntegrationArgs = Parameters<NonNullable<AnkiOptions['createAnkiIntegration']>>[0];

function createFakeTracker(targetMinimized = false) {
  return {
    onGeometryChange: null as ((geometry: WindowGeometry) => void) | null,
    onWindowFound: null as ((geometry: WindowGeometry) => void) | null,
    onWindowLost: null as (() => void) | null,
    onWindowFocusChange: null as ((focused: boolean) => void) | null,
    isTargetWindowMinimized: () => targetMinimized,
    started: false,
    start() {
      this.started = true;
    },
  };
}

function createAnkiHarness(overrides: Partial<AnkiOptions> = {}) {
  const state = {
    created: [] as CreateAnkiIntegrationArgs[],
    started: 0,
    stored: [] as unknown[],
  };
  const options: AnkiOptions = {
    getResolvedConfig: () => ({ ankiConnect: { enabled: true } as never }),
    getSubtitleTimingTracker: () => ({}),
    getMpvClient: () => ({ send: () => {} }),
    getRuntimeOptionsManager: () => ({
      getEffectiveAnkiConnectConfig: (config) => config as never,
    }),
    createAnkiIntegration: (args) => {
      state.created.push(args);
      return {
        start: () => {
          state.started += 1;
        },
      };
    },
    setAnkiIntegration: (integration) => {
      state.stored.push(integration);
    },
    showDesktopNotification: () => {},
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: false,
    }),
    getKnownWordCacheStatePath: () => '/tmp/known-words-cache.json',
    ...overrides,
  };
  return { options, state };
}

/**
 * Runs initializeOverlayRuntime with Anki disabled and returns the fake tracker plus a
 * snapshot of what the tracker callbacks did since initialization finished.
 */
function createTrackerHarness(config: { overlayVisible: boolean; targetMinimized?: boolean }) {
  const tracker = createFakeTracker(config.targetMinimized);
  const fresh = () => ({
    bounds: [] as WindowGeometry[],
    visibilityRefreshes: 0,
    subtitleRefreshes: 0,
    shortcutSyncs: 0,
    hiddenWindows: [] as string[],
  });
  let state = fresh();
  const overlayWindows = ['visible', 'modal'].map((name) => ({
    hide: () => {
      state.hiddenWindows.push(name);
    },
  }));

  const options: RuntimeOptions = {
    ...createAnkiHarness({
      getResolvedConfig: () => ({ ankiConnect: { enabled: false } as never }),
    }).options,
    backendOverride: null,
    createMainWindow: () => {},
    registerGlobalShortcuts: () => {},
    updateVisibleOverlayBounds: (geometry) => {
      state.bounds.push(geometry);
    },
    isVisibleOverlayVisible: () => config.overlayVisible,
    updateVisibleOverlayVisibility: () => {
      state.visibilityRefreshes += 1;
    },
    refreshCurrentSubtitle: () => {
      state.subtitleRefreshes += 1;
    },
    getOverlayWindows: () => overlayWindows as never,
    syncOverlayShortcuts: () => {
      state.shortcutSyncs += 1;
    },
    setWindowTracker: () => {},
    getMpvSocketPath: () => '/tmp/mpv.sock',
    createWindowTracker: () => tracker as never,
  };

  initializeOverlayRuntime(options);
  state = fresh();

  return { tracker, snapshot: () => state };
}

test('startOverlayWindowTracker starts tracker for the current mpv socket', () => {
  const tracker = createFakeTracker();
  const created: Array<[string | null | undefined, string | null | undefined]> = [];
  const stored: unknown[] = [];
  const bounds: WindowGeometry[] = [];
  let visibilityRefreshes = 0;
  let handlersInstalledAtStart = false;
  tracker.start = () => {
    handlersInstalledAtStart = tracker.onWindowFound !== null && tracker.onWindowLost !== null;
    tracker.started = true;
  };

  const result = startOverlayWindowTracker({
    backendOverride: 'windows',
    getMpvSocketPath: () => '\\\\.\\pipe\\subminer-socket',
    createWindowTracker: (override, socketPath) => {
      created.push([override, socketPath]);
      return tracker as never;
    },
    setWindowTracker: (nextTracker) => {
      stored.push(nextTracker);
    },
    updateVisibleOverlayBounds: (geometry) => {
      bounds.push(geometry);
    },
    isVisibleOverlayVisible: () => true,
    updateVisibleOverlayVisibility: () => {
      visibilityRefreshes += 1;
    },
    getOverlayWindows: () => [],
    syncOverlayShortcuts: () => {},
  });

  const geometry = { x: 10, y: 20, width: 300, height: 200 };
  tracker.onWindowFound?.(geometry);

  assert.equal(result, tracker);
  assert.deepEqual(created, [['windows', '\\\\.\\pipe\\subminer-socket']]);
  assert.deepEqual(stored, [tracker]);
  assert.equal(tracker.started, true);
  assert.equal(handlersInstalledAtStart, true);
  assert.deepEqual(bounds, [geometry]);
  assert.equal(visibilityRefreshes, 1);
});

const ankiNotCreatedCases: Array<{ name: string; overrides: Partial<AnkiOptions> }> = [
  {
    name: 'ankiConnect is disabled',
    overrides: { getResolvedConfig: () => ({ ankiConnect: { enabled: false } as never }) },
  },
  { name: 'an integration already exists', overrides: { getAnkiIntegration: () => ({}) } },
];

for (const c of ankiNotCreatedCases) {
  test(`initializeOverlayAnkiIntegration returns false when ${c.name}`, () => {
    const { options, state } = createAnkiHarness(c.overrides);

    assert.equal(initializeOverlayAnkiIntegration(options), false);
    assert.equal(state.created.length, 0);
    assert.equal(state.started, 0);
    assert.equal(state.stored.length, 0);
  });
}

test('initializeOverlayAnkiIntegration creates, starts, and stores the integration when enabled', () => {
  const { options, state } = createAnkiHarness();

  assert.equal(initializeOverlayAnkiIntegration(options), true);
  assert.equal(state.created.length, 1);
  assert.equal(state.created[0]!.config.enabled, true);
  assert.equal(state.started, 1);
  assert.equal(state.stored.length, 1);
});

test('initializeOverlayAnkiIntegration can skip starting the Anki integration transport', () => {
  const { options, state } = createAnkiHarness({ shouldStartAnkiIntegration: () => false });

  initializeOverlayAnkiIntegration(options);

  assert.equal(state.created.length, 1);
  assert.equal(state.started, 0);
  assert.equal(state.stored.length, 1);
});

test('initializeOverlayAnkiIntegration merges shared ai config with Anki overrides', () => {
  const { options, state } = createAnkiHarness({
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
  });

  initializeOverlayAnkiIntegration(options);

  const aiConfig = state.created[0]!.aiConfig;
  assert.equal(aiConfig.apiKey, 'shared-key');
  assert.equal(aiConfig.baseUrl, 'https://openrouter.ai/api');
  assert.equal(aiConfig.model, 'openrouter/anki-model');
  assert.equal(aiConfig.systemPrompt, 'Translate mined sentence text.');
});

test('initializeOverlayRuntime initializes the Anki integration after the window tracker starts', () => {
  const { options: ankiOptions, state } = createAnkiHarness();
  const tracker = createFakeTracker();

  initializeOverlayRuntime({
    ...ankiOptions,
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
    createWindowTracker: () => tracker as never,
  });

  assert.equal(tracker.started, true);
  assert.equal(state.created.length, 1);
  assert.equal(state.started, 1);
  assert.equal(state.stored.length, 1);
});

const trackerEventCases: Array<{
  name: string;
  overlayVisible: boolean;
  targetMinimized?: boolean;
  trigger: (tracker: ReturnType<typeof createFakeTracker>) => void;
  expected: Partial<ReturnType<ReturnType<typeof createTrackerHarness>['snapshot']>>;
}> = [
  {
    name: 'refreshes the visible overlay and re-syncs shortcuts when focus changes while shown',
    overlayVisible: true,
    trigger: (tracker) => tracker.onWindowFocusChange?.(true),
    expected: { visibilityRefreshes: 1, shortcutSyncs: 1 },
  },
  {
    name: 'only re-syncs shortcuts when focus changes while the overlay is hidden',
    overlayVisible: false,
    trigger: (tracker) => tracker.onWindowFocusChange?.(true),
    expected: { shortcutSyncs: 1 },
  },
  {
    name: 'restores bounds, visibility, and subtitle when the tracker finds the target window again',
    overlayVisible: true,
    trigger: (tracker) => tracker.onWindowFound?.({ x: 100, y: 200, width: 1280, height: 720 }),
    expected: {
      bounds: [{ x: 100, y: 200, width: 1280, height: 720 }],
      visibilityRefreshes: 1,
      subtitleRefreshes: 1,
    },
  },
  {
    name: 'hides overlay windows when the tracker loses a minimized target window',
    overlayVisible: true,
    targetMinimized: true,
    trigger: (tracker) => tracker.onWindowLost?.(),
    expected: { hiddenWindows: ['visible', 'modal'], shortcutSyncs: 1 },
  },
  {
    name: 'refreshes visibility instead of hiding when the tracker loses a non-minimized target',
    overlayVisible: true,
    targetMinimized: false,
    trigger: (tracker) => tracker.onWindowLost?.(),
    expected: { visibilityRefreshes: 1 },
  },
];

for (const c of trackerEventCases) {
  test(`initializeOverlayRuntime ${c.name}`, () => {
    const { tracker, snapshot } = createTrackerHarness(c);

    c.trigger(tracker);

    assert.deepEqual(snapshot(), {
      bounds: [],
      visibilityRefreshes: 0,
      subtitleRefreshes: 0,
      shortcutSyncs: 0,
      hiddenWindows: [],
      ...c.expected,
    });
  });
}
