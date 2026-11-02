import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG } from '../../config/definitions';
import type { ResolvedConfig } from '../../types';
import { createAppState } from '../state';
import { createStartupAssembly, type StartupAssemblyDeps } from './startup';

type MpvClient = NonNullable<ReturnType<typeof createAppState>['mpvClient']>;

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function lazy(name: string) {
  return (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
}

/** Builds deps whose every function throws unless a test overrides it. */
function makeDeps(overrides: Partial<StartupAssemblyDeps> = {}): StartupAssemblyDeps {
  const thrower = <T extends object>(name: string): T =>
    new Proxy({}, { get: (_target, key: string) => lazy(`${name}.${key}`) }) as T;
  return {
    appState: createAppState({ mpvSocketPath: '/tmp/subminer-startup.sock', texthookerPort: 0 }),
    configService: thrower('configService'),
    logger: thrower('logger'),
    logConfigWarning: lazy('logConfigWarning'),
    configDir: '/tmp/subminer-startup-assembly-test',
    userDataPath: '/tmp/subminer-startup-assembly-test',
    defaultTexthookerPort: 5174,
    showErrorBox: lazy('showErrorBox'),
    reportFatalError: lazy('reportFatalError'),
    requestAppQuit: lazy('requestAppQuit'),
    logging: {
      setLogLevel: lazy('setLogLevel'),
      setLogRotation: lazy('setLogRotation'),
      setLogFileToggles: lazy('setLogFileToggles'),
    },
    overlayManager: thrower('overlayManager'),
    overlayModalRuntime: thrower('overlayModalRuntime'),
    subtitleWsService: thrower('subtitleWsService'),
    annotationSubtitleWsService: thrower('annotationSubtitleWsService'),
    texthookerService: thrower('texthookerService'),
    configDerivedRuntime: thrower('configDerivedRuntime'),
    hasMpvWebsocketPlugin: lazy('hasMpvWebsocketPlugin'),
    loadSubtitlePosition: lazy('loadSubtitlePosition'),
    setSecondarySubMode: lazy('setSecondarySubMode'),
    syncOverlayMpvSubtitleSuppression: lazy('syncOverlayMpvSubtitleSuppression'),
    yomitan: thrower('yomitan'),
    stats: { ensureImmersionTrackerStarted: lazy('ensureImmersionTrackerStarted') } as never,
    firstRunSetup: thrower('firstRunSetup'),
    youtube: thrower('youtube'),
    overlay: thrower('overlay'),
    shortcuts: thrower('shortcuts'),
    settings: { hotReload: thrower('settings.hotReload') },
    mining: thrower('mining'),
    getAnilist: lazy('getAnilist'),
    getJellyfin: lazy('getJellyfin'),
    getPlayback: lazy('getPlayback'),
    getPluginHost: lazy('getPluginHost'),
    getOverlayNotifications: lazy('getOverlayNotifications'),
    getOverlayVisibilityComposer: lazy('getOverlayVisibilityComposer'),
    getVisibleOverlayInteraction: lazy('getVisibleOverlayInteraction'),
    getOverlayGeometry: lazy('getOverlayGeometry'),
    handleInitialArgs: lazy('handleInitialArgs'),
    ...overrides,
  };
}

/** Deps for the overlay prerequisites, recording what each step does. */
function makePrereqDeps(calls: string[]): Partial<StartupAssemblyDeps> {
  const record = (name: string) => () => {
    calls.push(name);
  };
  return {
    configService: { getConfig: () => DEFAULT_CONFIG } as StartupAssemblyDeps['configService'],
    loadSubtitlePosition: record('loadSubtitlePosition'),
    shortcuts: {
      refreshCurrentSessionBindings: record('refreshCurrentSessionBindings'),
      refreshOverlayShortcuts: record('refreshOverlayShortcuts'),
    } as unknown as StartupAssemblyDeps['shortcuts'],
    getPlayback: () =>
      ({
        createMpvClient: () => {
          calls.push('createMpvClient');
          return {} as MpvClient;
        },
        subtitleProcessingController: {
          invalidateTokenizationCache: record('invalidateTokenizationCache'),
        },
        getSubtitlePrefetchService: () => ({ onSeek: record('prefetch.onSeek') }),
        getLastObservedTimePos: () => 12,
      }) as unknown as ReturnType<StartupAssemblyDeps['getPlayback']>,
    getOverlayVisibilityComposer: () => ({
      broadcastRuntimeOptionsChanged: record('broadcastRuntimeOptionsChanged'),
    }),
  };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createStartupAssembly does not call its deps while constructing', () => {
  const startup = createStartupAssembly(makeDeps());
  assert.equal(typeof startup.runAppReady, 'function');
});

test('overlay startup prerequisites are created once', () => {
  const calls: string[] = [];
  const deps = makeDeps(makePrereqDeps(calls));
  const startup = createStartupAssembly(deps);

  startup.ensureOverlayStartupPrereqs();
  deps.appState.subtitlePosition = { yPercent: 10 };
  startup.ensureOverlayStartupPrereqs();

  assert.deepEqual(calls, [
    'loadSubtitlePosition',
    'refreshCurrentSessionBindings',
    'createMpvClient',
    'refreshCurrentSessionBindings',
  ]);
  assert.ok(deps.appState.runtimeOptionsManager);
  assert.ok(deps.appState.subtitleTimingTracker);
});

test('a runtime option change refreshes tokenization, prefetch, overlay state and shortcuts', () => {
  const calls: string[] = [];
  const deps = makeDeps(makePrereqDeps(calls));
  createStartupAssembly(deps).ensureOverlayStartupPrereqs();
  calls.length = 0;

  const manager = deps.appState.runtimeOptionsManager;
  assert.ok(manager);
  manager.setOptionValue('subtitle.annotation.jlpt', true);

  assert.deepEqual(calls, [
    'invalidateTokenizationCache',
    'prefetch.onSeek',
    'broadcastRuntimeOptionsChanged',
    'refreshOverlayShortcuts',
  ]);
});

test('a headless launch without --refresh-known-words runs the initial CLI command', async () => {
  let handled = 0;
  const deps = makeDeps({
    handleInitialArgs: () => {
      handled += 1;
    },
  });

  await createStartupAssembly(deps).runHeadlessInitialCommand();

  assert.equal(handled, 1);
});

test('a headless known-word refresh fails and quits when AnkiConnect is disabled', async (t) => {
  const previousExitCode = process.exitCode;
  t.after(() => {
    process.exitCode = previousExitCode ?? 0;
  });
  const config: ResolvedConfig = {
    ...DEFAULT_CONFIG,
    ankiConnect: { ...DEFAULT_CONFIG.ankiConnect, enabled: false },
  };
  const errors: string[] = [];
  let quits = 0;
  const deps = makeDeps({
    configService: { getConfig: () => config } as StartupAssemblyDeps['configService'],
    logger: { error: (message: string) => errors.push(message) } as never,
    requestAppQuit: () => {
      quits += 1;
    },
  });
  deps.appState.initialArgs = { refreshKnownWords: true } as never;

  await createStartupAssembly(deps).runHeadlessInitialCommand();

  assert.deepEqual(errors, [
    'Headless known-word refresh failed: AnkiConnect integration not enabled',
  ]);
  assert.equal(process.exitCode, 1);
  assert.equal(quits, 1);
});
