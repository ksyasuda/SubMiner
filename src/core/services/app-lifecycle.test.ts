import assert from 'node:assert/strict';
import test from 'node:test';
import { CliArgs } from '../../cli/args';
import { AppLifecycleServiceDeps, startAppLifecycle } from './app-lifecycle';

function makeArgs(overrides: Partial<CliArgs> = {}): CliArgs {
  return {
    background: false,
    managedPlayback: false,
    start: false,
    launchMpv: false,
    launchMpvTargets: [],
    stop: false,
    toggle: false,
    toggleVisibleOverlay: false,
    togglePrimarySubtitleBar: false,
    yomitan: false,
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
    markWatched: false,
    toggleSubtitleSidebar: false,
    openRuntimeOptions: false,
    openSessionHelp: false,
    openControllerSelect: false,
    openControllerDebug: false,
    openJimaku: false,
    openYoutubePicker: false,
    openPlaylistBrowser: false,
    openCharacterDictionary: false,
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
    dictionaryCandidates: false,
    dictionarySelect: false,
    dictionaryAnilistId: undefined,
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
    texthookerOpenBrowser: false,
    help: false,
    appPing: false,
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
    exitApp: (code) => {
      calls.push(`exit:${code}`);
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

test('startAppLifecycle app ping exits non-zero immediately when no running instance owns the lock', () => {
  const { deps, calls, getLockCalls } = createDeps({
    shouldStartApp: () => false,
  });

  startAppLifecycle(makeArgs({ appPing: true }), deps);

  assert.equal(getLockCalls(), 1);
  assert.deepEqual(calls, ['exit:1']);
});

test('startAppLifecycle app ping exits zero immediately when another instance owns the lock', () => {
  let lockCalls = 0;
  const { deps, calls } = createDeps({
    shouldStartApp: () => false,
    requestSingleInstanceLock: () => {
      lockCalls += 1;
      return false;
    },
  });

  startAppLifecycle(makeArgs({ appPing: true }), deps);

  assert.equal(lockCalls, 1);
  assert.deepEqual(calls, ['exit:0']);
});

test('startAppLifecycle queues second-instance commands until app ready runtime completes', async () => {
  const handled: string[] = [];
  let secondInstanceHandler: ((_event: unknown, argv: string[]) => void) | null = null;
  let readyHandler: (() => Promise<void>) | null = null;
  let releaseReady: (() => void) | null = null;
  const readyFinished = new Promise<void>((resolve) => {
    releaseReady = resolve;
  });

  const { deps } = createDeps({
    shouldStartApp: () => true,
    onSecondInstance: (handler) => {
      secondInstanceHandler = handler;
    },
    parseArgs: (argv) => makeArgs({ start: argv.includes('--start') }),
    handleCliCommand: (args, source) => {
      handled.push(`${source}:${args.start ? 'start' : 'other'}`);
    },
    whenReady: (handler) => {
      readyHandler = handler;
    },
    onReady: async () => {
      await readyFinished;
      handled.push('ready');
    },
  });

  startAppLifecycle(makeArgs({ background: true }), deps);

  const runSecondInstance = (argv: string[]) => {
    assert.ok(secondInstanceHandler);
    (secondInstanceHandler as (_event: unknown, argv: string[]) => void)({}, argv);
  };
  const runReady = () => {
    assert.ok(readyHandler);
    return (readyHandler as () => Promise<void>)();
  };

  runSecondInstance(['SubMiner', '--start']);
  assert.deepEqual(handled, []);

  const readyRun = runReady();
  await Promise.resolve();
  assert.deepEqual(handled, []);

  assert.ok(releaseReady);
  (releaseReady as () => void)();
  await readyRun;

  assert.deepEqual(handled, ['ready', 'second-instance:start']);

  runSecondInstance(['SubMiner', '--start']);
  assert.deepEqual(handled, ['ready', 'second-instance:start', 'second-instance:start']);
});

test('startAppLifecycle routes control socket commands through the second-instance queue', async () => {
  const handled: string[] = [];
  let controlArgvHandler: ((argv: string[]) => void) | null = null;
  let readyHandler: (() => Promise<void>) | null = null;
  let releaseReady: (() => void) | null = null;
  const readyFinished = new Promise<void>((resolve) => {
    releaseReady = resolve;
  });

  const { deps } = createDeps({
    shouldStartApp: () => true,
    parseArgs: (argv) => makeArgs({ start: argv.includes('--start') }),
    handleCliCommand: (args, source) => {
      handled.push(`${source}:${args.start ? 'start' : 'other'}`);
    },
    startControlServer: (handler) => {
      controlArgvHandler = handler;
      return () => {
        handled.push('control-close');
      };
    },
    whenReady: (handler) => {
      readyHandler = handler;
    },
    onReady: async () => {
      await readyFinished;
      handled.push('ready');
    },
  });

  let willQuitHandler: (() => void) | null = null;
  deps.onWillQuit = (handler) => {
    willQuitHandler = handler;
  };

  startAppLifecycle(makeArgs({ background: true }), deps);

  assert.ok(controlArgvHandler);
  (controlArgvHandler as (argv: string[]) => void)(['--start']);
  assert.deepEqual(handled, []);

  assert.ok(readyHandler);
  const readyRun = (readyHandler as () => Promise<void>)();
  await Promise.resolve();
  assert.deepEqual(handled, []);

  assert.ok(releaseReady);
  (releaseReady as () => void)();
  await readyRun;
  assert.deepEqual(handled, ['ready', 'second-instance:start']);

  assert.ok(willQuitHandler);
  (willQuitHandler as () => void)();
  assert.deepEqual(handled, ['ready', 'second-instance:start', 'control-close']);
});

test('startAppLifecycle drains queued second-instance commands when app ready runtime fails', async () => {
  const handled: string[] = [];
  let controlArgvHandler: ((argv: string[]) => void) | null = null;
  let readyHandler: (() => Promise<void>) | null = null;

  const { deps } = createDeps({
    shouldStartApp: () => true,
    parseArgs: (argv) => makeArgs({ start: argv.includes('--start') }),
    handleCliCommand: (args, source) => {
      handled.push(`${source}:${args.start ? 'start' : 'other'}`);
    },
    startControlServer: (handler) => {
      controlArgvHandler = handler;
    },
    whenReady: (handler) => {
      readyHandler = handler;
    },
    onReady: async () => {
      handled.push('ready');
      throw new Error('ready failed');
    },
  });

  startAppLifecycle(makeArgs({ background: true }), deps);

  assert.ok(controlArgvHandler);
  (controlArgvHandler as (argv: string[]) => void)(['--start']);
  assert.deepEqual(handled, []);

  assert.ok(readyHandler);
  await assert.rejects((readyHandler as () => Promise<void>)(), /ready failed/);

  assert.deepEqual(handled, ['ready', 'second-instance:start']);

  (controlArgvHandler as (argv: string[]) => void)(['--start']);
  assert.deepEqual(handled, ['ready', 'second-instance:start', 'second-instance:start']);
});

test('startAppLifecycle quits macOS config-only launch when all windows close', () => {
  let windowAllClosedHandler: (() => void) | null = null;
  const { deps, calls } = createDeps({
    shouldStartApp: () => true,
    isDarwinPlatform: () => true,
    shouldQuitOnWindowAllClosed: () => true,
    onWindowAllClosed: (handler) => {
      windowAllClosedHandler = handler;
    },
  });

  startAppLifecycle(makeArgs({ settings: true }), deps);

  const handler = windowAllClosedHandler as (() => void) | null;
  assert.ok(handler);
  handler();
  assert.deepEqual(calls, ['quitApp']);
});

test('startAppLifecycle quits macOS setup-only launch when all windows close', () => {
  let windowAllClosedHandler: (() => void) | null = null;
  const { deps, calls } = createDeps({
    shouldStartApp: () => true,
    isDarwinPlatform: () => true,
    shouldQuitOnWindowAllClosed: () => true,
    onWindowAllClosed: (handler) => {
      windowAllClosedHandler = handler;
    },
  });

  startAppLifecycle(makeArgs({ setup: true }), deps);

  const handler = windowAllClosedHandler as (() => void) | null;
  assert.ok(handler);
  handler();
  assert.deepEqual(calls, ['quitApp']);
});
