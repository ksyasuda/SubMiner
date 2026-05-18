import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import type { LauncherCommandContext } from './context.js';
import { runPlaybackCommandWithDeps } from './playback-command.js';
import { state } from '../mpv.js';

function createContext(): LauncherCommandContext {
  return {
    args: {
      backend: 'auto',
      directory: '.',
      recursive: false,
      profile: '',
      startOverlay: false,
      youtubeMode: 'download',
      whisperBin: '',
      whisperModel: '',
      whisperVadModel: '',
      whisperThreads: 0,
      youtubeSubgenOutDir: '',
      youtubeSubgenAudioFormat: '',
      youtubeSubgenKeepTemp: false,
      youtubeFixWithAi: false,
      youtubePrimarySubLangs: [],
      youtubeSecondarySubLangs: [],
      youtubeAudioLangs: [],
      youtubeWhisperSourceLanguage: '',
      aiConfig: {},
      useTexthooker: false,
      autoStartOverlay: false,
      texthookerOnly: false,
      texthookerOpenBrowser: false,
      useRofi: false,
      logLevel: 'info',
      passwordStore: '',
      target: 'https://www.youtube.com/watch?v=65Ovd7t8sNw',
      targetKind: 'url',
      jimakuApiKey: '',
      jimakuApiKeyCommand: '',
      jimakuApiBaseUrl: '',
      jimakuLanguagePreference: 'ja',
      jimakuMaxEntryResults: 20,
      jellyfin: false,
      jellyfinLogin: false,
      jellyfinLogout: false,
      jellyfinPlay: false,
      jellyfinDiscovery: false,
      dictionary: false,
      dictionaryCandidates: false,
      dictionarySelect: false,
      stats: false,
      doctor: false,
      doctorRefreshKnownWords: false,
      version: false,
      configSettings: false,
      configPath: false,
      configShow: false,
      mpvIdle: false,
      mpvSocket: false,
      mpvStatus: false,
      mpvArgs: '',
      appPassthrough: false,
      appArgs: [],
      jellyfinServer: '',
      jellyfinUsername: '',
      jellyfinPassword: '',
      launchMode: 'normal',
    },
    scriptPath: '/tmp/subminer',
    scriptName: 'subminer',
    mpvSocketPath: '/tmp/subminer.sock',
    pluginRuntimeConfig: {
      socketPath: '/tmp/subminer.sock',
      binaryPath: '',
      backend: 'auto',
      autoStart: true,
      autoStartVisibleOverlay: true,
      autoStartPauseUntilReady: true,
      texthookerEnabled: false,
      aniskipEnabled: true,
      aniskipButtonKey: 'TAB',
    },
    appPath: '/tmp/SubMiner.AppImage',
    launcherJellyfinConfig: {},
    processAdapter: {
      platform: () => 'linux',
      onSignal: () => {},
      writeStdout: () => {},
      exit: (_code: number): never => {
        throw new Error('unexpected exit');
      },
      setExitCode: () => {},
    },
  };
}

test('youtube playback launches overlay with app-owned youtube flow args', async () => {
  const calls: string[] = [];
  const context = createContext();
  context.pluginRuntimeConfig = {
    ...context.pluginRuntimeConfig,
    autoStart: false,
    autoStartVisibleOverlay: false,
    autoStartPauseUntilReady: false,
  };
  const receivedStartMpvOptions: Record<string, unknown>[] = [];

  await runPlaybackCommandWithDeps(context, {
    ensurePlaybackSetupReady: async () => {},
    chooseTarget: async (_args, _scriptPath) => ({ target: context.args.target, kind: 'url' }),
    checkDependencies: () => {},
    registerCleanup: () => {},
    startMpv: async (
      _target,
      _targetKind,
      _args,
      _socketPath,
      _appPath,
      _preloadedSubtitles,
      options,
    ) => {
      if (options) {
        receivedStartMpvOptions.push(options as Record<string, unknown>);
      }
      calls.push('startMpv');
    },
    waitForUnixSocketReady: async () => true,
    startOverlay: async (_appPath, _args, _socketPath, extraAppArgs = []) => {
      calls.push(`startOverlay:${extraAppArgs.join(' ')}`);
    },
    launchAppCommandDetached: (_appPath: string, appArgs: string[]) => {
      calls.push(`launch:${appArgs.join(' ')}`);
    },
    log: () => {},
    cleanupPlaybackSession: async () => {},
    getMpvProc: () => null,
  });

  assert.deepEqual(calls, [
    'startMpv',
    'startOverlay:--youtube-play https://www.youtube.com/watch?v=65Ovd7t8sNw --youtube-mode download',
  ]);
  assert.equal(receivedStartMpvOptions[0]?.startPaused, true);
  assert.equal(receivedStartMpvOptions[0]?.disableYoutubeSubtitleAutoLoad, true);
});

test('plugin auto-start playback leaves app lifetime to managed-playback owner', async () => {
  const context = createContext();
  context.args = {
    ...context.args,
    target: '/tmp/movie.mkv',
    targetKind: 'file',
  };
  context.pluginRuntimeConfig = {
    socketPath: '/tmp/subminer.sock',
    binaryPath: '',
    backend: 'auto',
    autoStart: true,
    autoStartVisibleOverlay: false,
    autoStartPauseUntilReady: false,
    texthookerEnabled: false,
    aniskipEnabled: true,
    aniskipButtonKey: 'TAB',
  };
  const appPath = context.appPath ?? '';
  state.appPath = appPath;
  state.overlayManagedByLauncher = false;
  const mpvProc = new EventEmitter() as EventEmitter & {
    exitCode: number | null;
    killed: boolean;
    kill: () => boolean;
  };
  mpvProc.exitCode = null;
  mpvProc.killed = false;
  mpvProc.kill = () => true;
  let cleanupSawManagedOverlay = true;

  try {
    await runPlaybackCommandWithDeps(context, {
      ensurePlaybackSetupReady: async () => {},
      chooseTarget: async () => ({ target: context.args.target, kind: 'file' }),
      checkDependencies: () => {},
      registerCleanup: () => {},
      startMpv: async () => {
        setTimeout(() => {
          mpvProc.exitCode = 0;
          mpvProc.emit('exit', 0);
        }, 5);
      },
      waitForUnixSocketReady: async () => true,
      startOverlay: async () => {
        throw new Error('startOverlay should not run when plugin auto-start is used');
      },
      launchAppCommandDetached: () => {},
      log: () => {},
      cleanupPlaybackSession: async () => {
        cleanupSawManagedOverlay = state.overlayManagedByLauncher;
      },
      getMpvProc: () => mpvProc as NonNullable<typeof state.mpvProc>,
    });

    assert.equal(cleanupSawManagedOverlay, false);
  } finally {
    state.appPath = '';
    state.overlayManagedByLauncher = false;
  }
});
