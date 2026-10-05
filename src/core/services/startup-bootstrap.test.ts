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
    yomitan: false,
    settings: false,
    syncWindow: false,
    youtubeBrowser: false,
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
    openTsukihime: false,
    openYoutubePicker: false,
    openPlaylistBrowser: false,
    replayCurrentSubtitle: false,
    playNextSubtitle: false,
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

function runBootstrap(
  argsOverrides: Partial<CliArgs> = {},
  depOverrides: Partial<Parameters<typeof runStartupBootstrapRuntime>[0]> = {},
) {
  const calls: string[] = [];
  const args = makeArgs(argsOverrides);
  const result = runStartupBootstrapRuntime({
    argv: ['node', 'main.ts'],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push('forceX11'),
    enforceUnsupportedWaylandMode: () => calls.push('enforceWayland'),
    getDefaultSocketPath: () => '/tmp/default.sock',
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push('startLifecycle'),
    ...depOverrides,
  });
  return { args, result, calls };
}

test('runStartupBootstrapRuntime maps CLI args to startup state', () => {
  const { args, result } = runBootstrap({
    socketPath: '/tmp/custom.sock',
    texthookerPort: 9001,
    backend: 'x11',
    autoStartOverlay: true,
    texthooker: true,
    background: true,
  });

  assert.equal(result.initialArgs, args);
  assert.equal(result.mpvSocketPath, '/tmp/custom.sock');
  assert.equal(result.texthookerPort, 9001);
  assert.equal(result.backendOverride, 'x11');
  assert.equal(result.autoStartOverlay, true);
  assert.equal(result.texthookerOnlyMode, true);
  assert.equal(result.backgroundMode, true);
});

test('runStartupBootstrapRuntime falls back to defaults when args are unset', () => {
  const { result } = runBootstrap();

  assert.equal(result.mpvSocketPath, '/tmp/default.sock');
  assert.equal(result.texthookerPort, 5174);
  assert.equal(result.backendOverride, null);
  assert.equal(result.autoStartOverlay, false);
  assert.equal(result.texthookerOnlyMode, false);
  assert.equal(result.backgroundMode, false);
});

const logLevelCases: Array<{ name: string; args: Partial<CliArgs>; expected: string[] }> = [
  {
    name: '--log-level applies the requested level',
    args: { logLevel: 'debug' },
    expected: ['setLog:debug:cli'],
  },
  { name: '--update defaults to warn', args: { update: true }, expected: ['setLog:warn:cli'] },
  {
    name: '--log-level wins over --update',
    args: { update: true, logLevel: 'info' },
    expected: ['setLog:info:cli'],
  },
  { name: '--debug leaves log verbosity alone', args: { debug: true }, expected: [] },
  { name: '--background leaves log level to config', args: { background: true }, expected: [] },
];

for (const c of logLevelCases) {
  test(`runStartupBootstrapRuntime log level: ${c.name}`, () => {
    const { calls } = runBootstrap(c.args);

    assert.deepEqual(
      calls.filter((call) => call.startsWith('setLog:')),
      c.expected,
    );
  });
}

test('runStartupBootstrapRuntime applies log level and platform setup before starting lifecycle', () => {
  const { calls } = runBootstrap({ logLevel: 'debug' });

  const index = (call: string) => calls.indexOf(call);
  assert.ok(index('setLog:debug:cli') >= 0);
  assert.ok(index('setLog:debug:cli') < index('forceX11'));
  assert.ok(index('forceX11') < index('startLifecycle'));
  assert.ok(index('enforceWayland') < index('startLifecycle'));
});

test('runStartupBootstrapRuntime skips lifecycle when generate-config flow handled', () => {
  const { calls } = runBootstrap({ generateConfig: true }, { runGenerateConfigFlow: () => true });

  assert.equal(calls.includes('startLifecycle'), false);
});
