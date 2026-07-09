import test from 'node:test';
import assert from 'node:assert/strict';
import type { Args } from './types.js';
import { runJellyfinPlayMenuWithDeps } from './jellyfin.js';

function createArgs(): Args {
  return {
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
    history: false,
    sync: false,
    syncHost: '',
    syncSnapshotPath: '',
    syncMergePath: '',
    syncRemoteCmd: '',
    syncDbPath: '',
    syncForce: false,
    logLevel: 'info',
    logRotation: 7,
    passwordStore: '',
    target: '',
    targetKind: '',
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
    logsExport: false,
    version: false,
    settings: false,
    configPath: false,
    configShow: false,
    mpvIdle: false,
    mpvSocket: false,
    mpvStatus: false,
    mpvArgs: '',
    appPassthrough: false,
    appArgs: [],
    jellyfinServer: 'https://jellyfin.example.test',
    jellyfinUsername: '',
    jellyfinPassword: '',
    launchMode: 'normal',
  };
}

test('Jellyfin playback ensures Linux runtime plugin before detached idle mpv bootstrap', async () => {
  const originalAccessToken = process.env.SUBMINER_JELLYFIN_ACCESS_TOKEN;
  const originalUserId = process.env.SUBMINER_JELLYFIN_USER_ID;
  process.env.SUBMINER_JELLYFIN_ACCESS_TOKEN = 'token';
  process.env.SUBMINER_JELLYFIN_USER_ID = 'user';
  const calls: string[] = [];

  try {
    await assert.rejects(
      () =>
        runJellyfinPlayMenuWithDeps(
          '/tmp/SubMiner.AppImage',
          createArgs(),
          '/tmp/subminer',
          '/tmp/subminer.sock',
          {
            loadLauncherJellyfinConfig: () => ({}),
            findRofiTheme: () => null,
            resolveJellyfinSelection: async () => 'item-123',
            resolveJellyfinSelectionViaApp: async () => {
              throw new Error('unexpected app-based selection');
            },
            hasStoredJellyfinSession: () => true,
            requestJellyfinPreviewAuthFromApp: async () => null,
            resolveLauncherMainConfigPath: () => '/tmp/SubMiner/config.jsonc',
            pathExists: () => false,
            ensureRuntimePluginReady: async () => {
              calls.push('plugin');
            },
            waitForUnixSocketReady: async () => {
              calls.push('wait');
              return true;
            },
            launchMpvIdleDetached: async () => {
              calls.push('launch');
            },
            resolveLauncherRuntimePluginPath: () => {
              calls.push('resolve');
              return '/tmp/plugin/main.lua';
            },
            runAppCommandWithInheritLogged: () => {
              calls.push('handoff');
              throw new Error('stop after handoff');
            },
            log: () => {},
          },
        ),
      /stop after handoff/,
    );

    assert.deepEqual(calls, ['plugin', 'resolve', 'launch', 'wait', 'handoff']);
  } finally {
    if (originalAccessToken === undefined) {
      delete process.env.SUBMINER_JELLYFIN_ACCESS_TOKEN;
    } else {
      process.env.SUBMINER_JELLYFIN_ACCESS_TOKEN = originalAccessToken;
    }
    if (originalUserId === undefined) {
      delete process.env.SUBMINER_JELLYFIN_USER_ID;
    } else {
      process.env.SUBMINER_JELLYFIN_USER_ID = originalUserId;
    }
  }
});
