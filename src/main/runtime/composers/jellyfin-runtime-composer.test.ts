import assert from 'node:assert/strict';
import test from 'node:test';
import {
  composeJellyfinRuntimeHandlers,
  createRestartJellyfinRemoteSessionAfterSetupLoginHandler,
} from './jellyfin-runtime-composer';

test('setup login restart uses auto-connect path without an active remote session', async () => {
  const startOptions: Array<{ explicit?: boolean } | undefined> = [];
  const restart = createRestartJellyfinRemoteSessionAfterSetupLoginHandler({
    getCurrentSession: () => null,
    startJellyfinRemoteSession: async (options) => {
      startOptions.push(options);
    },
  });

  await restart();

  assert.deepEqual(startOptions, [undefined]);
});

test('setup login restart explicitly refreshes an active remote session', async () => {
  const startOptions: Array<{ explicit?: boolean } | undefined> = [];
  const restart = createRestartJellyfinRemoteSessionAfterSetupLoginHandler({
    getCurrentSession: () => ({ stop: () => {} }),
    startJellyfinRemoteSession: async (options) => {
      startOptions.push(options);
    },
  });

  await restart();

  assert.deepEqual(startOptions, [{ explicit: true }]);
});

test('composeJellyfinRuntimeHandlers returns callable jellyfin runtime handlers', () => {
  let activePlayback: unknown = null;
  let lastProgressAtMs = 0;
  const composed = composeJellyfinRuntimeHandlers({
    getResolvedJellyfinConfigMainDeps: {
      getResolvedConfig: () => ({ jellyfin: { enabled: false, serverUrl: '' } }) as never,
      loadStoredSession: () => null,
      getEnv: () => undefined,
    },
    getJellyfinClientInfoMainDeps: {
      getResolvedJellyfinConfig: () => ({}) as never,
      getHostName: () => 'workstation',
      defaultClientName: 'SubMiner',
      defaultClientVersion: 'test',
    },
    waitForMpvConnectedMainDeps: {
      getMpvClient: () => null,
      now: () => Date.now(),
      sleep: async () => {},
    },
    launchMpvIdleForJellyfinPlaybackMainDeps: {
      getSocketPath: () => '/tmp/test-mpv.sock',
      getLaunchMode: () => 'normal',
      platform: 'linux',
      execPath: process.execPath,
      getDefaultMpvLogPath: () => '/tmp/test-mpv.log',
      defaultMpvArgs: [],
      removeSocketPath: () => {},
      spawnMpv: () => ({ unref: () => {} }) as never,
      logWarn: () => {},
      logInfo: () => {},
    },
    ensureMpvConnectedForJellyfinPlaybackMainDeps: {
      getMpvClient: () => null,
      setMpvClient: () => {},
      createMpvClient: () => ({}) as never,
      getAutoLaunchInFlight: () => null,
      setAutoLaunchInFlight: () => {},
      connectTimeoutMs: 10,
      autoLaunchTimeoutMs: 10,
    },
    preloadJellyfinExternalSubtitlesMainDeps: {
      listJellyfinSubtitleTracks: async () => [],
      getMpvClient: () => null,
      sendMpvCommand: () => {},
      wait: async () => {},
      cacheSubtitleTrack: async () => ({ path: '/tmp/sub.srt', cleanupDir: '/tmp/subs' }),
      cleanupCachedSubtitles: () => {},
      logDebug: () => {},
    },
    playJellyfinItemInMpvMainDeps: {
      getMpvClient: () => null,
      resolvePlaybackPlan: async () => ({
        mode: 'direct',
        url: 'https://example.test/video.m3u8',
        title: 'Episode 1',
        itemTitle: 'Episode 1',
        seriesTitle: null,
        seasonNumber: null,
        episodeNumber: null,
        startTimeTicks: 0,
        audioStreamIndex: null,
        subtitleStreamIndex: null,
      }),
      applyJellyfinMpvDefaults: () => {},
      showVisibleOverlay: () => {},
      sendMpvCommand: () => {},
      armQuitOnDisconnect: () => {},
      schedule: () => undefined,
      convertTicksToSeconds: () => 0,
      setActivePlayback: (value) => {
        activePlayback = value;
      },
      setLastProgressAtMs: (value) => {
        lastProgressAtMs = value;
      },
      reportPlaying: () => {},
      showMpvOsd: () => {},
    },
    remoteComposerOptions: {
      getConfiguredSession: () => null,
      logWarn: () => {},
      getMpvClient: () => null,
      sendMpvCommand: () => {},
      jellyfinTicksToSeconds: () => 0,
      getActivePlayback: () => activePlayback as never,
      clearActivePlayback: () => {
        activePlayback = null;
      },
      getSession: () => null,
      getNow: () => Date.now(),
      getLastProgressAtMs: () => lastProgressAtMs,
      setLastProgressAtMs: (value) => {
        lastProgressAtMs = value;
      },
      progressIntervalMs: 3000,
      ticksPerSecond: 10_000_000,
      logDebug: () => {},
    },
    handleJellyfinAuthCommandsMainDeps: {
      patchRawConfig: () => {},
      authenticateWithPassword: async () => ({
        serverUrl: 'https://example.test',
        username: 'user',
        accessToken: 'token',
        userId: 'id',
      }),
      saveStoredSession: () => {},
      clearStoredSession: () => {},
      logInfo: () => {},
    },
    handleJellyfinListCommandsMainDeps: {
      listJellyfinLibraries: async () => [],
      listJellyfinItems: async () => [],
      listJellyfinSubtitleTracks: async () => [],
      writeJellyfinPreviewAuth: () => {},
      logInfo: () => {},
    },
    handleJellyfinPlayCommandMainDeps: {
      logWarn: () => {},
    },
    handleJellyfinRemoteAnnounceCommandMainDeps: {
      getRemoteSession: () => null,
      logInfo: () => {},
      logWarn: () => {},
    },
    startJellyfinRemoteSessionMainDeps: {
      getCurrentSession: () => null,
      setCurrentSession: () => {},
      createRemoteSessionService: () =>
        ({
          start: async () => {},
        }) as never,
      defaultDeviceId: 'dev',
      defaultClientName: 'SubMiner',
      defaultClientVersion: 'test',
      getHostName: () => 'workstation',
      logInfo: () => {},
      logWarn: () => {},
    },
    stopJellyfinRemoteSessionMainDeps: {
      getCurrentSession: () => null,
      setCurrentSession: () => {},
      clearActivePlayback: () => {
        activePlayback = null;
      },
    },
    runJellyfinCommandMainDeps: {
      defaultServerUrl: 'https://example.test',
    },
    maybeFocusExistingJellyfinSetupWindowMainDeps: {
      getSetupWindow: () => null,
    },
    openJellyfinSetupWindowMainDeps: {
      createSetupWindow: () =>
        ({
          focus: () => {},
          webContents: { on: () => {} },
          loadURL: () => {},
          on: () => {},
          isDestroyed: () => false,
          close: () => {},
        }) as never,
      buildSetupFormHtml: (state) => `<html>${state.selectedServerUrl}${state.username}</html>`,
      parseSubmissionUrl: () => null,
      authenticateWithPassword: async () => ({
        serverUrl: 'https://example.test',
        username: 'user',
        accessToken: 'token',
        userId: 'id',
      }),
      saveStoredSession: () => {},
      clearStoredSession: () => {},
      patchJellyfinConfig: () => {},
      logInfo: () => {},
      logError: () => {},
      showMpvOsd: () => {},
      clearSetupWindow: () => {},
      setSetupWindow: () => {},
      encodeURIComponent,
      defaultServerUrl: 'https://example.test',
      hasStoredSession: () => false,
    },
  });

  assert.equal(typeof composed.getResolvedJellyfinConfig, 'function');
  assert.equal(typeof composed.getJellyfinClientInfo, 'function');
  assert.equal(typeof composed.reportJellyfinRemoteProgress, 'function');
  assert.equal(typeof composed.reportJellyfinRemoteStopped, 'function');
  assert.equal(typeof composed.handleJellyfinRemotePlay, 'function');
  assert.equal(typeof composed.handleJellyfinRemotePlaystate, 'function');
  assert.equal(typeof composed.handleJellyfinRemoteGeneralCommand, 'function');
  assert.equal(typeof composed.playJellyfinItemInMpv, 'function');
  assert.equal(typeof composed.cleanupJellyfinSubtitleCache, 'function');
  assert.equal(typeof composed.startJellyfinRemoteSession, 'function');
  assert.equal(typeof composed.stopJellyfinRemoteSession, 'function');
  assert.equal(typeof composed.runJellyfinCommand, 'function');
  assert.equal(typeof composed.openJellyfinSetupWindow, 'function');

  // getResolvedJellyfinConfig forwards to the injected getResolvedConfig dep
  const jellyfinConfig = composed.getResolvedJellyfinConfig();
  assert.equal(jellyfinConfig.enabled, false);
  assert.equal(jellyfinConfig.serverUrl, '');
});
