import assert from 'node:assert/strict';
import test from 'node:test';
import { runAppReadyRuntime, type AppReadyRuntimeDeps } from './startup';

// Every dep records its name in `calls`; overrides replace individual deps.
function createAppReadyDeps(overrides: Partial<AppReadyRuntimeDeps> = {}) {
  const calls: string[] = [];
  const record = (name: string) => () => {
    calls.push(name);
  };
  const recordAsync = (name: string) => async () => {
    calls.push(name);
  };
  const deps: AppReadyRuntimeDeps = {
    ensureDefaultConfigBootstrap: record('bootstrap'),
    loadSubtitlePosition: record('load-subtitle-position'),
    resolveKeybindings: record('resolve-keybindings'),
    createMpvClient: record('create-mpv'),
    reloadConfig: record('reload-config'),
    getResolvedConfig: () => ({
      websocket: { enabled: false },
      annotationWebsocket: { enabled: false },
      texthooker: { launchAtStartup: false },
    }),
    getConfigWarnings: () => [],
    logConfigWarning: record('config-warning'),
    setLogLevel: record('set-log-level'),
    initRuntimeOptionsManager: record('init-runtime-options'),
    setSecondarySubMode: record('set-secondary-sub-mode'),
    defaultSecondarySubMode: 'hover',
    defaultWebsocketPort: 0,
    defaultAnnotationWebsocketPort: 0,
    defaultTexthookerPort: 0,
    hasMpvWebsocketPlugin: () => false,
    startSubtitleWebsocket: record('subtitle-ws'),
    startAnnotationWebsocket: record('annotation-ws'),
    startTexthooker: record('texthooker'),
    log: record('log'),
    createMecabTokenizerAndCheck: recordAsync('mecab'),
    createSubtitleTimingTracker: record('subtitle-timing'),
    createImmersionTracker: record('immersion'),
    startJellyfinRemoteSession: recordAsync('jellyfin'),
    loadYomitanExtension: recordAsync('load-yomitan'),
    handleFirstRunSetup: recordAsync('first-run'),
    prewarmSubtitleDictionaries: recordAsync('prewarm'),
    startBackgroundWarmups: record('warmups'),
    texthookerOnlyMode: false,
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
    setVisibleOverlayVisible: record('visible-overlay'),
    initializeOverlayRuntime: record('init-overlay'),
    handleInitialArgs: record('handle-initial-args'),
    shouldUseMinimalStartup: () => false,
    shouldSkipHeavyStartup: () => false,
    ...overrides,
  };
  return { calls, deps };
}

function assertBefore(calls: string[], earlier: string, later: string) {
  const earlierIndex = calls.indexOf(earlier);
  const laterIndex = calls.indexOf(later);
  assert.notEqual(earlierIndex, -1, `${earlier} was not called`);
  assert.notEqual(laterIndex, -1, `${later} was not called`);
  assert.ok(earlierIndex < laterIndex, `${earlier} should run before ${later}`);
}

test('runAppReadyRuntime minimal startup skips Yomitan and first-run setup while still handling CLI args', async () => {
  const { calls, deps } = createAppReadyDeps({ shouldUseMinimalStartup: () => true });

  await runAppReadyRuntime(deps);

  assert.ok(calls.includes('handle-initial-args'));
  for (const skipped of ['load-yomitan', 'first-run', 'create-mpv', 'init-overlay', 'warmups']) {
    assert.equal(calls.includes(skipped), false, `${skipped} should be skipped`);
  }
});

test('runAppReadyRuntime headless refresh bootstraps Anki runtime without UI startup', async () => {
  const { calls, deps } = createAppReadyDeps({
    shouldRunHeadlessInitialCommand: () => true,
    runHeadlessInitialCommand: async () => {
      calls.push('run-headless-command');
    },
  });

  await runAppReadyRuntime(deps);

  assertBefore(calls, 'init-runtime-options', 'run-headless-command');
  for (const skipped of ['create-mpv', 'load-yomitan', 'init-overlay', 'handle-initial-args']) {
    assert.equal(calls.includes(skipped), false, `${skipped} should be skipped`);
  }
});

test('runAppReadyRuntime loads Yomitan before headless overlay fallback initialization', async () => {
  const { calls, deps } = createAppReadyDeps({ shouldRunHeadlessInitialCommand: () => true });

  await runAppReadyRuntime(deps);

  assertBefore(calls, 'create-mpv', 'subtitle-timing');
  assertBefore(calls, 'load-yomitan', 'init-overlay');
  assertBefore(calls, 'init-overlay', 'handle-initial-args');
});

test('runAppReadyRuntime auto-initializes overlay runtime before warmups and Yomitan', async () => {
  const { calls, deps } = createAppReadyDeps({
    shouldAutoInitializeOverlayRuntimeFromConfig: () => true,
  });

  await runAppReadyRuntime(deps);

  assertBefore(calls, 'init-overlay', 'warmups');
  assert.equal(calls.includes('load-yomitan'), false);
});

test('runAppReadyRuntime reuses guarded Yomitan loader after scheduling startup warmups', async () => {
  const { calls, deps } = createAppReadyDeps({
    ensureYomitanExtensionLoaded: async () => {
      calls.push('load-yomitan-guarded');
    },
  });

  await runAppReadyRuntime(deps);

  assert.equal(calls.includes('load-yomitan'), false);
  assertBefore(calls, 'warmups', 'load-yomitan-guarded');
});
