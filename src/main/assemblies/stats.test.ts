import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG } from '../../config/definitions';
import type { ResolvedConfig } from '../../types';
import type { AppState } from '../state';
import { createStatsAssembly, type StatsAssemblyDeps } from './stats';

const trackingDisabledConfig: ResolvedConfig = {
  ...DEFAULT_CONFIG,
  immersionTracking: { ...DEFAULT_CONFIG.immersionTracking, enabled: false },
};

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(overrides: Partial<StatsAssemblyDeps> = {}): StatsAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: {
      immersionTracker: null,
      statsServer: null,
      statsStartupInProgress: false,
      mpvClient: null,
      mpvSocketPath: '/tmp/subminer-stats-test.sock',
      mecabTokenizer: null,
      ankiIntegration: null,
      currentMediaPath: null,
      currentMediaTitle: null,
    },
    getConfig: lazy('getConfig'),
    logger: { info: lazy('info'), warn: lazy('warn'), error: lazy('error'), debug: lazy('debug') },
    userDataPath: '/tmp/subminer-stats-assembly-test',
    mainDirname: '/tmp/subminer-stats-assembly-test/dist',
    defaultImmersionDbPath: '/tmp/subminer-stats-assembly-test/immersion.sqlite',
    anilistRateLimiter: { acquire: lazy('acquire'), recordResponse: lazy('recordResponse') },
    statsWindow: { registerToggle: lazy('registerToggle'), toggle: lazy('toggle') },
    yomitanParserRuntimeDeps: {
      getYomitanExt: lazy('getYomitanExt'),
      getYomitanSession: lazy('getYomitanSession'),
      getYomitanParserWindow: lazy('getYomitanParserWindow'),
      setYomitanParserWindow: lazy('setYomitanParserWindow'),
      getYomitanParserReadyPromise: lazy('getYomitanParserReadyPromise'),
      setYomitanParserReadyPromise: lazy('setYomitanParserReadyPromise'),
      getYomitanParserInitPromise: lazy('getYomitanParserInitPromise'),
      setYomitanParserInitPromise: lazy('setYomitanParserInitPromise'),
    },
    getYomitanAnkiDeckName: lazy('getYomitanAnkiDeckName'),
    getTokenizeSubtitle: lazy('getTokenizeSubtitle'),
    createMecabTokenizerAndCheck: lazy('createMecabTokenizerAndCheck'),
    getOverlayNotifications: lazy('getOverlayNotifications'),
    getOverlayBounds: lazy('getOverlayBounds'),
    onStatsOverlayVisibilityChanged: lazy('onStatsOverlayVisibilityChanged'),
    runAnilistPostWatchUpdate: lazy('runAnilistPostWatchUpdate'),
    openExternal: lazy('openExternal'),
    requestAppQuit: lazy('requestAppQuit'),
    ...overrides,
  };
}

function withTracker(
  deps: StatsAssemblyDeps,
  markActiveVideoWatched: () => Promise<boolean>,
): StatsAssemblyDeps {
  const tracker = { markActiveVideoWatched } as unknown as NonNullable<
    AppState['immersionTracker']
  >;
  return { ...deps, appState: { ...deps.appState, immersionTracker: tracker } };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createStatsAssembly does not call its deps while constructing', () => {
  const stats = createStatsAssembly(makeDeps());
  assert.equal(typeof stats.ensureImmersionTrackerStarted, 'function');
  assert.equal(typeof stats.toggleStatsOverlay, 'function');
});

test('a manual watched mark tries to start the immersion tracker first', async () => {
  const infoLogs: string[] = [];
  const stats = createStatsAssembly(
    makeDeps({
      getConfig: () => trackingDisabledConfig,
      logger: {
        info: (message: string) => {
          infoLogs.push(message);
        },
        warn: () => {},
        error: () => {},
        debug: () => {},
      },
    }),
  );

  // With tracking disabled, the startup attempt logs and leaves no tracker to mark.
  assert.equal(await stats.markActiveVideoWatched(), false);
  assert.deepEqual(infoLogs, ['Immersion tracking disabled in config']);
});

for (const c of [
  { name: 'runs the AniList update after a successful mark', marked: true, anilistRuns: 1 },
  { name: 'skips the AniList update when nothing was marked', marked: false, anilistRuns: 0 },
]) {
  test(`a manual watched mark ${c.name}`, async () => {
    let anilistRuns = 0;
    const stats = createStatsAssembly(
      withTracker(
        makeDeps({
          runAnilistPostWatchUpdate: async () => {
            anilistRuns += 1;
          },
        }),
        async () => c.marked,
      ),
    );

    assert.equal(await stats.markActiveVideoWatched(), c.marked);
    assert.equal(anilistRuns, c.anilistRuns);
  });
}

test('a failed AniList update does not undo a manual watched mark', async () => {
  const warnings: string[] = [];
  const stats = createStatsAssembly(
    withTracker(
      makeDeps({
        logger: {
          info: () => {},
          warn: (message: string) => {
            warnings.push(message);
          },
          error: () => {},
          debug: () => {},
        },
        runAnilistPostWatchUpdate: async () => {
          throw new Error('AniList unavailable');
        },
      }),
      async () => true,
    ),
  );

  assert.equal(await stats.markActiveVideoWatched(), true);
  assert.deepEqual(warnings, [
    'Failed to run AniList post-watch update after manual watched mark:',
  ]);
});
