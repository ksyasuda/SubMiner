import test from 'node:test';
import assert from 'node:assert/strict';
import type { LauncherCommandContext } from './context.js';
import { runPlaybackCommandWithDeps } from './playback-command.js';

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
      stats: false,
      doctor: false,
      doctorRefreshKnownWords: false,
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
      autoStart: true,
      autoStartVisibleOverlay: true,
      autoStartPauseUntilReady: true,
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
  let receivedStartMpvOptions: Record<string, unknown> | null = null;

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
      receivedStartMpvOptions = options ?? null;
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
  assert.deepEqual(receivedStartMpvOptions, {
    startPaused: true,
    disableYoutubeSubtitleAutoLoad: true,
  });
});
