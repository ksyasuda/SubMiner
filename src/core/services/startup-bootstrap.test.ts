import test from 'node:test';
import assert from 'node:assert/strict';
import { runStartupBootstrapRuntime } from './startup';
import { CliArgs } from '../../cli/args';

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
    settings: false,
    configSettings: false,
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
    autoStartOverlay: false,
    generateConfig: false,
    backupOverwrite: false,
    debug: false,
    ...overrides,
  };
}

test('runStartupBootstrapRuntime configures startup state and starts lifecycle', () => {
  const calls: string[] = [];
  const args = makeArgs({
    logLevel: 'debug',
    socketPath: '/tmp/custom.sock',
    texthookerPort: 9001,
    backend: 'x11',
    autoStartOverlay: true,
    texthooker: true,
  });

  const result = runStartupBootstrapRuntime({
    argv: ['node', 'main.ts', '--log-level', 'debug'],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push('forceX11'),
    enforceUnsupportedWaylandMode: () => calls.push('enforceWayland'),
    getDefaultSocketPath: () => '/tmp/default.sock',
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push('startLifecycle'),
  });

  assert.equal(result.initialArgs, args);
  assert.equal(result.mpvSocketPath, '/tmp/custom.sock');
  assert.equal(result.texthookerPort, 9001);
  assert.equal(result.backendOverride, 'x11');
  assert.equal(result.autoStartOverlay, true);
  assert.equal(result.texthookerOnlyMode, true);
  assert.equal(result.backgroundMode, false);
  assert.deepEqual(calls, ['setLog:debug:cli', 'forceX11', 'enforceWayland', 'startLifecycle']);
});

test('runStartupBootstrapRuntime keeps log-level precedence for repeated calls', () => {
  const calls: string[] = [];
  const args = makeArgs({
    logLevel: 'warn',
  });

  runStartupBootstrapRuntime({
    argv: ['node', 'main.ts', '--log-level', 'warn'],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push('forceX11'),
    enforceUnsupportedWaylandMode: () => calls.push('enforceWayland'),
    getDefaultSocketPath: () => '/tmp/default.sock',
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push('startLifecycle'),
  });

  assert.deepEqual(calls.slice(0, 3), ['setLog:warn:cli', 'forceX11', 'enforceWayland']);
});

test('runStartupBootstrapRuntime remains lifecycle-stable with Jellyfin CLI flags', () => {
  const calls: string[] = [];
  const args = makeArgs({
    jellyfin: true,
    jellyfinLibraries: true,
    socketPath: '/tmp/stable.sock',
    texthookerPort: 8888,
  });

  const result = runStartupBootstrapRuntime({
    argv: ['node', 'main.ts', '--jellyfin', '--jellyfin-libraries'],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push('forceX11'),
    enforceUnsupportedWaylandMode: () => calls.push('enforceWayland'),
    getDefaultSocketPath: () => '/tmp/default.sock',
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push('startLifecycle'),
  });

  assert.equal(result.mpvSocketPath, '/tmp/stable.sock');
  assert.equal(result.texthookerPort, 8888);
  assert.equal(result.backendOverride, null);
  assert.equal(result.autoStartOverlay, false);
  assert.equal(result.texthookerOnlyMode, false);
  assert.equal(result.backgroundMode, false);
  assert.deepEqual(calls, ['forceX11', 'enforceWayland', 'startLifecycle']);
});

test('runStartupBootstrapRuntime keeps --debug separate from log verbosity', () => {
  const calls: string[] = [];
  const args = makeArgs({
    debug: true,
  });

  runStartupBootstrapRuntime({
    argv: ['node', 'main.ts', '--debug'],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push('forceX11'),
    enforceUnsupportedWaylandMode: () => calls.push('enforceWayland'),
    getDefaultSocketPath: () => '/tmp/default.sock',
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push('startLifecycle'),
  });

  assert.deepEqual(calls, ['forceX11', 'enforceWayland', 'startLifecycle']);
});

test('runStartupBootstrapRuntime skips lifecycle when generate-config flow handled', () => {
  const calls: string[] = [];
  const args = makeArgs({ generateConfig: true, logLevel: 'warn' });

  const result = runStartupBootstrapRuntime({
    argv: ['node', 'main.ts', '--generate-config'],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push('forceX11'),
    enforceUnsupportedWaylandMode: () => calls.push('enforceWayland'),
    getDefaultSocketPath: () => '/tmp/default.sock',
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => true,
    startAppLifecycle: () => calls.push('startLifecycle'),
  });

  assert.equal(result.mpvSocketPath, '/tmp/default.sock');
  assert.equal(result.texthookerPort, 5174);
  assert.equal(result.backendOverride, null);
  assert.equal(result.backgroundMode, false);
  assert.deepEqual(calls, ['setLog:warn:cli', 'forceX11', 'enforceWayland']);
});

test('runStartupBootstrapRuntime enables quiet background mode by default', () => {
  const calls: string[] = [];
  const args = makeArgs({ background: true });

  const result = runStartupBootstrapRuntime({
    argv: ['node', 'main.ts', '--background'],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push('forceX11'),
    enforceUnsupportedWaylandMode: () => calls.push('enforceWayland'),
    getDefaultSocketPath: () => '/tmp/default.sock',
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push('startLifecycle'),
  });

  assert.equal(result.backgroundMode, true);
  assert.deepEqual(calls, ['setLog:warn:cli', 'forceX11', 'enforceWayland', 'startLifecycle']);
});

test('runStartupBootstrapRuntime enables quiet update mode by default', () => {
  const calls: string[] = [];
  const args = makeArgs({ update: true });

  const result = runStartupBootstrapRuntime({
    argv: ['node', 'main.ts', '--update'],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push('forceX11'),
    enforceUnsupportedWaylandMode: () => calls.push('enforceWayland'),
    getDefaultSocketPath: () => '/tmp/default.sock',
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push('startLifecycle'),
  });

  assert.equal(result.backgroundMode, false);
  assert.deepEqual(calls, ['setLog:warn:cli', 'forceX11', 'enforceWayland', 'startLifecycle']);
});
