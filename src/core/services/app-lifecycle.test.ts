import assert from 'node:assert/strict';
import test from 'node:test';
import { CliArgs } from '../../cli/args';
import { AppLifecycleServiceDeps, startAppLifecycle } from './app-lifecycle';

function makeArgs(overrides: Partial<CliArgs> = {}): CliArgs {
  return {
    background: false,
    start: false,
    launchMpv: false,
    launchMpvTargets: [],
    stop: false,
    toggle: false,
    toggleVisibleOverlay: false,
    settings: false,
    setup: false,
    show: false,
    hide: false,
    showVisibleOverlay: false,
    hideVisibleOverlay: false,
    copySubtitle: false,
    copySubtitleMultiple: false,
    mineSentence: false,
    mineSentenceMultiple: false,
    updateLastCardFromClipboard: false,
    refreshKnownWords: false,
    toggleSecondarySub: false,
    triggerFieldGrouping: false,
    triggerSubsync: false,
    markAudioCard: false,
    toggleStatsOverlay: false,
    toggleSubtitleSidebar: false,
    openRuntimeOptions: false,
    openSessionHelp: false,
    openControllerSelect: false,
    openControllerDebug: false,
    openJimaku: false,
    openYoutubePicker: false,
    openPlaylistBrowser: false,
    replayCurrentSubtitle: false,
    playNextSubtitle: false,
    shiftSubDelayPrevLine: false,
    shiftSubDelayNextLine: false,
    cycleRuntimeOptionId: undefined,
    cycleRuntimeOptionDirection: undefined,
    anilistStatus: false,
    anilistLogout: false,
    anilistSetup: false,
    anilistRetryQueue: false,
    dictionary: false,
    stats: false,
    jellyfin: false,
    jellyfinLogin: false,
    jellyfinLogout: false,
    jellyfinLibraries: false,
    jellyfinItems: false,
    jellyfinSubtitles: false,
    jellyfinSubtitleUrlsOnly: false,
    jellyfinPlay: false,
    jellyfinRemoteAnnounce: false,
    jellyfinPreviewAuth: false,
    texthooker: false,
    help: false,
    autoStartOverlay: false,
    generateConfig: false,
    backupOverwrite: false,
    debug: false,
    ...overrides,
  };
}

function createDeps(overrides: Partial<AppLifecycleServiceDeps> = {}) {
  const calls: string[] = [];
  let lockCalls = 0;

  const deps: AppLifecycleServiceDeps = {
    shouldStartApp: () => false,
    parseArgs: () => makeArgs(),
    requestSingleInstanceLock: () => {
      lockCalls += 1;
      return true;
    },
    quitApp: () => {
      calls.push('quitApp');
    },
    onSecondInstance: () => {},
    handleCliCommand: () => {},
    printHelp: () => {
      calls.push('printHelp');
    },
    logNoRunningInstance: () => {
      calls.push('logNoRunningInstance');
    },
    whenReady: () => {},
    onWindowAllClosed: () => {},
    onWillQuit: () => {},
    onActivate: () => {},
    isDarwinPlatform: () => false,
    onReady: async () => {},
    onWillQuitCleanup: () => {},
    shouldRestoreWindowsOnActivate: () => false,
    restoreWindowsOnActivate: () => {},
    shouldQuitOnWindowAllClosed: () => true,
    ...overrides,
  };

  return { deps, calls, getLockCalls: () => lockCalls };
}

test('startAppLifecycle handles --help without acquiring single-instance lock', () => {
  const { deps, calls, getLockCalls } = createDeps({
    shouldStartApp: () => false,
  });

  startAppLifecycle(makeArgs({ help: true }), deps);

  assert.equal(getLockCalls(), 0);
  assert.deepEqual(calls, ['printHelp', 'quitApp']);
});

test('startAppLifecycle still acquires lock for startup commands', () => {
  const { deps, getLockCalls } = createDeps({
    shouldStartApp: () => true,
    whenReady: () => {},
  });

  startAppLifecycle(makeArgs({ start: true }), deps);

  assert.equal(getLockCalls(), 1);
});
