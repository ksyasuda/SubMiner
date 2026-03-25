import test from 'node:test';
import assert from 'node:assert/strict';
import { CliArgs } from '../../cli/args';
import { CliCommandServiceDeps, handleCliCommand } from './cli-command';

function makeArgs(overrides: Partial<CliArgs> = {}): CliArgs {
  return {
    background: false,
    start: false,
    launchMpv: false,
    launchMpvTargets: [],
    youtubePlay: undefined,
    youtubeMode: undefined,
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
    toggleSecondarySub: false,
    triggerFieldGrouping: false,
    triggerSubsync: false,
    markAudioCard: false,
    refreshKnownWords: false,
    openRuntimeOptions: false,
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

function createDeps(overrides: Partial<CliCommandServiceDeps> = {}) {
  const calls: string[] = [];
  let mpvSocketPath = '/tmp/subminer.sock';
  let texthookerPort = 5174;
  const osd: string[] = [];

  const deps: CliCommandServiceDeps = {
    getMpvSocketPath: () => mpvSocketPath,
    setMpvSocketPath: (socketPath) => {
      mpvSocketPath = socketPath;
      calls.push(`setMpvSocketPath:${socketPath}`);
    },
    setMpvClientSocketPath: (socketPath) => {
      calls.push(`setMpvClientSocketPath:${socketPath}`);
    },
    hasMpvClient: () => true,
    connectMpvClient: () => {
      calls.push('connectMpvClient');
    },
    isTexthookerRunning: () => false,
    setTexthookerPort: (port) => {
      texthookerPort = port;
      calls.push(`setTexthookerPort:${port}`);
    },
    getTexthookerPort: () => texthookerPort,
    shouldOpenTexthookerBrowser: () => true,
    ensureTexthookerRunning: (port) => {
      calls.push(`ensureTexthookerRunning:${port}`);
    },
    openTexthookerInBrowser: (url) => {
      calls.push(`openTexthookerInBrowser:${url}`);
    },
    stopApp: () => {
      calls.push('stopApp');
    },
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => {
      calls.push('initializeOverlayRuntime');
    },
    toggleVisibleOverlay: () => {
      calls.push('toggleVisibleOverlay');
    },
    openYomitanSettingsDelayed: (delayMs) => {
      calls.push(`openYomitanSettingsDelayed:${delayMs}`);
    },
    openFirstRunSetup: () => {
      calls.push('openFirstRunSetup');
    },
    setVisibleOverlayVisible: (visible) => {
      calls.push(`setVisibleOverlayVisible:${visible}`);
    },
    copyCurrentSubtitle: () => {
      calls.push('copyCurrentSubtitle');
    },
    startPendingMultiCopy: (timeoutMs) => {
      calls.push(`startPendingMultiCopy:${timeoutMs}`);
    },
    mineSentenceCard: async () => {
      calls.push('mineSentenceCard');
    },
    startPendingMineSentenceMultiple: (timeoutMs) => {
      calls.push(`startPendingMineSentenceMultiple:${timeoutMs}`);
    },
    updateLastCardFromClipboard: async () => {
      calls.push('updateLastCardFromClipboard');
    },
    refreshKnownWords: async () => {
      calls.push('refreshKnownWords');
    },
    cycleSecondarySubMode: () => {
      calls.push('cycleSecondarySubMode');
    },
    triggerFieldGrouping: async () => {
      calls.push('triggerFieldGrouping');
    },
    triggerSubsyncFromConfig: async () => {
      calls.push('triggerSubsyncFromConfig');
    },
    markLastCardAsAudioCard: async () => {
      calls.push('markLastCardAsAudioCard');
    },
    openRuntimeOptionsPalette: () => {
      calls.push('openRuntimeOptionsPalette');
    },
    getAnilistStatus: () => ({
      tokenStatus: 'resolved',
      tokenSource: 'stored',
      tokenMessage: null,
      tokenResolvedAt: 1,
      tokenErrorAt: null,
      queuePending: 2,
      queueReady: 1,
      queueDeadLetter: 0,
      queueLastAttemptAt: 2,
      queueLastError: null,
    }),
    clearAnilistToken: () => {
      calls.push('clearAnilistToken');
    },
    openAnilistSetup: () => {
      calls.push('openAnilistSetup');
    },
    openJellyfinSetup: () => {
      calls.push('openJellyfinSetup');
    },
    getAnilistQueueStatus: () => ({
      pending: 2,
      ready: 1,
      deadLetter: 0,
      lastAttemptAt: null,
      lastError: null,
    }),
    retryAnilistQueue: async () => {
      calls.push('retryAnilistQueue');
      return { ok: true, message: 'AniList retry processed.' };
    },
    generateCharacterDictionary: async () => ({
      zipPath: '/tmp/anilist-1.zip',
      fromCache: false,
      mediaId: 1,
      mediaTitle: 'Test',
      entryCount: 10,
    }),
    runStatsCommand: async () => {
      calls.push('runStatsCommand');
    },
    runJellyfinCommand: async () => {
      calls.push('runJellyfinCommand');
    },
    runYoutubePlaybackFlow: async (request) => {
      calls.push(`runYoutubePlaybackFlow:${request.url}:${request.mode}:${request.source}`);
    },
    printHelp: () => {
      calls.push('printHelp');
    },
    hasMainWindow: () => true,
    getMultiCopyTimeoutMs: () => 2500,
    showMpvOsd: (text) => {
      osd.push(text);
    },
    log: (message) => {
      calls.push(`log:${message}`);
    },
    warn: (message) => {
      calls.push(`warn:${message}`);
    },
    error: (message) => {
      calls.push(`error:${message}`);
    },
    ...overrides,
  };

  return { deps, calls, osd };
}

test('handleCliCommand starts youtube playback flow on initial launch', () => {
  const { deps, calls } = createDeps({
    runYoutubePlaybackFlow: async (request) => {
      calls.push(`youtube:${request.url}:${request.mode}`);
    },
  });

  handleCliCommand(
    makeArgs({ youtubePlay: 'https://youtube.com/watch?v=abc', youtubeMode: 'generate' }),
    'initial',
    deps,
  );

  assert.deepEqual(calls, [
    'initializeOverlayRuntime',
    'youtube:https://youtube.com/watch?v=abc:generate',
  ]);
});

test('handleCliCommand defaults youtube mode to download when omitted', () => {
  const { deps, calls } = createDeps({
    runYoutubePlaybackFlow: async (request) => {
      calls.push(`youtube:${request.url}:${request.mode}`);
    },
  });

  handleCliCommand(makeArgs({ youtubePlay: 'https://youtube.com/watch?v=abc' }), 'initial', deps);

  assert.deepEqual(calls, [
    'initializeOverlayRuntime',
    'youtube:https://youtube.com/watch?v=abc:download',
  ]);
});

test('handleCliCommand reuses initialized overlay runtime for second-instance youtube playback', () => {
  const { deps, calls } = createDeps({
    isOverlayRuntimeInitialized: () => true,
    runYoutubePlaybackFlow: async (request) => {
      calls.push(`youtube:${request.url}:${request.mode}:${request.source}`);
    },
  });

  handleCliCommand(
    makeArgs({ youtubePlay: 'https://youtube.com/watch?v=abc', youtubeMode: 'download' }),
    'second-instance',
    deps,
  );

  assert.deepEqual(calls, ['youtube:https://youtube.com/watch?v=abc:download:second-instance']);
});

test('handleCliCommand reports youtube playback flow failures to logs and OSD', async () => {
  const { deps, calls, osd } = createDeps({
    runYoutubePlaybackFlow: async () => {
      throw new Error('yt failed');
    },
  });

  handleCliCommand(
    makeArgs({ youtubePlay: 'https://youtube.com/watch?v=abc', youtubeMode: 'download' }),
    'initial',
    deps,
  );
  await new Promise((resolve) => setImmediate(resolve));

  assert.ok(calls.some((value) => value.startsWith('error:runYoutubePlaybackFlow failed:')));
  assert.ok(osd.includes('YouTube playback failed: yt failed'));
});

test('handleCliCommand reconnects MPV for second-instance --start when overlay runtime is already initialized', () => {
  const { deps, calls } = createDeps({
    isOverlayRuntimeInitialized: () => true,
  });
  const args = makeArgs({ start: true });

  handleCliCommand(args, 'second-instance', deps);

  assert.ok(calls.includes('setMpvClientSocketPath:/tmp/subminer.sock'));
  assert.equal(
    calls.some((value) => value.includes('connectMpvClient')),
    true,
  );
  assert.equal(
    calls.some((value) => value.includes('initializeOverlayRuntime')),
    false,
  );
});

test('handleCliCommand processes --start for second-instance when overlay runtime is not initialized', () => {
  const { deps, calls } = createDeps();
  const args = makeArgs({ start: true });

  handleCliCommand(args, 'second-instance', deps);

  assert.equal(
    calls.some((value) => value === 'log:Ignoring --start because SubMiner is already running.'),
    false,
  );
  assert.ok(calls.includes('setMpvClientSocketPath:/tmp/subminer.sock'));
  assert.equal(
    calls.some((value) => value.includes('connectMpvClient')),
    true,
  );
});

test('handleCliCommand opens first-run setup window for --setup', () => {
  const { deps, calls } = createDeps();

  handleCliCommand(makeArgs({ setup: true }), 'initial', deps);

  assert.ok(calls.includes('openFirstRunSetup'));
  assert.ok(calls.includes('log:Opened first-run setup flow.'));
  assert.equal(calls.includes('openYomitanSettingsDelayed:1000'), false);
});

test('handleCliCommand dispatches stats command without overlay startup', async () => {
  const { deps, calls } = createDeps({
    runStatsCommand: async () => {
      calls.push('runStatsCommand');
    },
  });

  handleCliCommand(makeArgs({ stats: true }), 'initial', deps);
  await Promise.resolve();

  assert.ok(calls.includes('runStatsCommand'));
  assert.equal(calls.includes('initializeOverlayRuntime'), false);
  assert.equal(calls.includes('connectMpvClient'), false);
});

test('handleCliCommand applies cli log level for second-instance commands', () => {
  const { deps, calls } = createDeps({
    setLogLevel: (level) => {
      calls.push(`setLogLevel:${level}`);
    },
  });

  handleCliCommand(makeArgs({ start: true, logLevel: 'debug' }), 'second-instance', deps);

  assert.ok(calls.includes('setLogLevel:debug'));
});

test('handleCliCommand runs texthooker flow with browser open', () => {
  const { deps, calls } = createDeps();
  const args = makeArgs({ texthooker: true });

  handleCliCommand(args, 'initial', deps);

  assert.ok(calls.includes('ensureTexthookerRunning:5174'));
  assert.ok(calls.includes('openTexthookerInBrowser:http://127.0.0.1:5174'));
});

test('handleCliCommand reports async mine errors to OSD', async () => {
  const { deps, calls, osd } = createDeps({
    mineSentenceCard: async () => {
      throw new Error('boom');
    },
  });

  handleCliCommand(makeArgs({ mineSentence: true }), 'initial', deps);
  await new Promise((resolve) => setImmediate(resolve));

  assert.ok(calls.some((value) => value.startsWith('error:mineSentenceCard failed:')));
  assert.ok(osd.some((value) => value.includes('Mine sentence failed: boom')));
});

test('handleCliCommand applies socket path and connects on start', () => {
  const { deps, calls } = createDeps();

  handleCliCommand(makeArgs({ start: true, socketPath: '/tmp/custom.sock' }), 'initial', deps);

  assert.ok(calls.includes('initializeOverlayRuntime'));
  assert.ok(calls.includes('setMpvSocketPath:/tmp/custom.sock'));
  assert.ok(calls.includes('setMpvClientSocketPath:/tmp/custom.sock'));
  assert.ok(calls.includes('connectMpvClient'));
});

test('handleCliCommand warns when texthooker port override used while running', () => {
  const { deps, calls } = createDeps({
    isTexthookerRunning: () => true,
  });

  handleCliCommand(makeArgs({ texthookerPort: 9999, texthooker: true }), 'initial', deps);

  assert.ok(
    calls.includes(
      'warn:Ignoring --port override because the texthooker server is already running.',
    ),
  );
  assert.equal(
    calls.some((value) => value === 'setTexthookerPort:9999'),
    false,
  );
});

test('handleCliCommand prints help and stops app when no window exists', () => {
  const { deps, calls } = createDeps({
    hasMainWindow: () => false,
  });

  handleCliCommand(makeArgs({ help: true }), 'initial', deps);

  assert.ok(calls.includes('printHelp'));
  assert.ok(calls.includes('stopApp'));
});

test('handleCliCommand reports async trigger-subsync errors to OSD', async () => {
  const { deps, calls, osd } = createDeps({
    triggerSubsyncFromConfig: async () => {
      throw new Error('subsync boom');
    },
  });

  handleCliCommand(makeArgs({ triggerSubsync: true }), 'initial', deps);
  await new Promise((resolve) => setImmediate(resolve));

  assert.ok(calls.some((value) => value.startsWith('error:triggerSubsyncFromConfig failed:')));
  assert.ok(osd.some((value) => value.includes('Subsync failed: subsync boom')));
});

test('handleCliCommand stops app for --stop command', () => {
  const { deps, calls } = createDeps();
  handleCliCommand(makeArgs({ stop: true }), 'initial', deps);
  assert.ok(calls.includes('log:Stopping SubMiner...'));
  assert.ok(calls.includes('stopApp'));
});

test('handleCliCommand still runs non-start actions on second-instance', () => {
  const { deps, calls } = createDeps();
  handleCliCommand(makeArgs({ start: true, toggleVisibleOverlay: true }), 'second-instance', deps);
  assert.ok(calls.includes('toggleVisibleOverlay'));
  assert.equal(
    calls.some((value) => value === 'connectMpvClient'),
    true,
  );
});

test('handleCliCommand connects MPV for toggle on second-instance', () => {
  const { deps, calls } = createDeps();
  handleCliCommand(makeArgs({ toggle: true }), 'second-instance', deps);
  assert.ok(calls.includes('toggleVisibleOverlay'));
  assert.equal(
    calls.some((value) => value === 'connectMpvClient'),
    true,
  );
});

test('handleCliCommand handles visibility and utility command dispatches', () => {
  const cases: Array<{
    args: Partial<CliArgs>;
    expected: string;
  }> = [
    { args: { settings: true }, expected: 'openYomitanSettingsDelayed:1000' },
    {
      args: { showVisibleOverlay: true },
      expected: 'setVisibleOverlayVisible:true',
    },
    {
      args: { hideVisibleOverlay: true },
      expected: 'setVisibleOverlayVisible:false',
    },
    { args: { copySubtitle: true }, expected: 'copyCurrentSubtitle' },
    {
      args: { copySubtitleMultiple: true },
      expected: 'startPendingMultiCopy:2500',
    },
    {
      args: { mineSentenceMultiple: true },
      expected: 'startPendingMineSentenceMultiple:2500',
    },
    { args: { toggleSecondarySub: true }, expected: 'cycleSecondarySubMode' },
    {
      args: { openRuntimeOptions: true },
      expected: 'openRuntimeOptionsPalette',
    },
    { args: { anilistLogout: true }, expected: 'clearAnilistToken' },
    { args: { anilistSetup: true }, expected: 'openAnilistSetup' },
    { args: { jellyfin: true }, expected: 'openJellyfinSetup' },
  ];

  for (const entry of cases) {
    const { deps, calls } = createDeps();
    handleCliCommand(makeArgs(entry.args), 'initial', deps);
    assert.ok(
      calls.includes(entry.expected),
      `expected call missing for args ${JSON.stringify(entry.args)}: ${entry.expected}`,
    );
  }
});

test('handleCliCommand logs AniList status details', () => {
  const { deps, calls } = createDeps();
  handleCliCommand(makeArgs({ anilistStatus: true }), 'initial', deps);
  assert.ok(calls.some((value) => value.startsWith('log:AniList token status:')));
  assert.ok(calls.some((value) => value.startsWith('log:AniList queue:')));
});

test('handleCliCommand runs AniList retry command', async () => {
  const { deps, calls } = createDeps();
  handleCliCommand(makeArgs({ anilistRetryQueue: true }), 'initial', deps);
  await new Promise((resolve) => setImmediate(resolve));
  assert.ok(calls.includes('retryAnilistQueue'));
  assert.ok(calls.includes('log:AniList retry processed.'));
});

test('handleCliCommand runs dictionary generation command', async () => {
  const { deps, calls } = createDeps({
    hasMainWindow: () => false,
    generateCharacterDictionary: async () => ({
      zipPath: '/tmp/anilist-9253.zip',
      fromCache: true,
      mediaId: 9253,
      mediaTitle: 'STEINS;GATE',
      entryCount: 314,
    }),
  });
  handleCliCommand(makeArgs({ dictionary: true }), 'initial', deps);
  await new Promise((resolve) => setImmediate(resolve));
  assert.ok(calls.includes('log:Generating character dictionary for current anime...'));
  assert.ok(
    calls.includes('log:Character dictionary cache hit: AniList 9253 (STEINS;GATE), entries=314'),
  );
  assert.ok(calls.includes('log:Dictionary ZIP: /tmp/anilist-9253.zip'));
  assert.ok(calls.includes('stopApp'));
});

test('handleCliCommand forwards --dictionary-target to dictionary runtime', async () => {
  let receivedTarget: string | undefined;
  const { deps } = createDeps({
    generateCharacterDictionary: async (targetPath?: string) => {
      receivedTarget = targetPath;
      return {
        zipPath: '/tmp/anilist-100.zip',
        fromCache: false,
        mediaId: 100,
        mediaTitle: 'Test',
        entryCount: 1,
      };
    },
  });

  handleCliCommand(
    makeArgs({ dictionary: true, dictionaryTarget: '/tmp/example-video.mkv' }),
    'initial',
    deps,
  );
  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(receivedTarget, '/tmp/example-video.mkv');
});

test('handleCliCommand does not dispatch runJellyfinCommand for non-Jellyfin commands', () => {
  const nonJellyfinArgs: Array<Partial<CliArgs>> = [
    { start: true },
    { copySubtitle: true },
    { toggleVisibleOverlay: true },
  ];

  for (const args of nonJellyfinArgs) {
    const { deps, calls } = createDeps();
    handleCliCommand(makeArgs(args), 'initial', deps);
    const runJellyfinCallCount = calls.filter((value) => value === 'runJellyfinCommand').length;
    assert.equal(
      runJellyfinCallCount,
      0,
      `Unexpected Jellyfin dispatch for args ${JSON.stringify(args)}`,
    );
  }
});

test('handleCliCommand runs jellyfin command dispatcher', async () => {
  const { deps, calls } = createDeps();
  handleCliCommand(makeArgs({ jellyfinLibraries: true }), 'initial', deps);
  handleCliCommand(makeArgs({ jellyfinSubtitles: true }), 'initial', deps);
  await new Promise((resolve) => setImmediate(resolve));
  const runJellyfinCallCount = calls.filter((value) => value === 'runJellyfinCommand').length;
  assert.equal(runJellyfinCallCount, 2);
});

test('handleCliCommand reports jellyfin command errors to OSD', async () => {
  const { deps, calls, osd } = createDeps({
    runJellyfinCommand: async () => {
      throw new Error('server offline');
    },
  });

  handleCliCommand(makeArgs({ jellyfinLibraries: true }), 'initial', deps);
  await new Promise((resolve) => setImmediate(resolve));

  assert.ok(calls.some((value) => value.startsWith('error:runJellyfinCommand failed:')));
  assert.ok(osd.some((value) => value.includes('Jellyfin command failed: server offline')));
});

test('handleCliCommand runs refresh-known-words command', () => {
  const { deps, calls } = createDeps();

  handleCliCommand(makeArgs({ refreshKnownWords: true }), 'initial', deps);

  assert.ok(calls.includes('refreshKnownWords'));
});

test('handleCliCommand stops app after headless initial refresh-known-words completes', async () => {
  const { deps, calls } = createDeps({
    hasMainWindow: () => false,
  });

  handleCliCommand(makeArgs({ refreshKnownWords: true }), 'initial', deps);
  await new Promise((resolve) => setImmediate(resolve));

  assert.ok(calls.includes('refreshKnownWords'));
  assert.ok(calls.includes('stopApp'));
});

test('handleCliCommand reports async refresh-known-words errors to OSD', async () => {
  const { deps, calls, osd } = createDeps({
    hasMainWindow: () => false,
    refreshKnownWords: async () => {
      throw new Error('refresh boom');
    },
  });

  handleCliCommand(makeArgs({ refreshKnownWords: true }), 'initial', deps);
  await new Promise((resolve) => setImmediate(resolve));

  assert.ok(calls.some((value) => value.startsWith('error:refreshKnownWords failed:')));
  assert.ok(osd.some((value) => value.includes('Refresh known words failed: refresh boom')));
  assert.ok(calls.includes('stopApp'));
});
