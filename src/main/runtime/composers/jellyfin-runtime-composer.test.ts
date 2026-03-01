import assert from 'node:assert/strict';
import test from 'node:test';
import { composeJellyfinRuntimeHandlers } from './jellyfin-runtime-composer';

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
      getDefaultJellyfinConfig: () => ({
        clientName: 'SubMiner',
        clientVersion: 'test',
        deviceId: 'dev',
      }),
    },
    waitForMpvConnectedMainDeps: {
      getMpvClient: () => null,
      now: () => Date.now(),
      sleep: async () => {},
    },
    launchMpvIdleForJellyfinPlaybackMainDeps: {
      getSocketPath: () => '/tmp/test-mpv.sock',
      platform: 'linux',
      execPath: process.execPath,
      defaultMpvLogPath: '/tmp/test-mpv.log',
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
      logDebug: () => {},
    },
    playJellyfinItemInMpvMainDeps: {
      getMpvClient: () => null,
      resolvePlaybackPlan: async () => ({
        mode: 'direct',
        url: 'https://example.test/video.m3u8',
        title: 'Episode 1',
        startTimeTicks: 0,
        audioStreamIndex: null,
        subtitleStreamIndex: null,
      }),
      applyJellyfinMpvDefaults: () => {},
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
      buildSetupFormHtml: (defaultServer, defaultUser) =>
        `<html>${defaultServer}${defaultUser}</html>`,
      parseSubmissionUrl: () => null,
      authenticateWithPassword: async () => ({
        serverUrl: 'https://example.test',
        username: 'user',
        accessToken: 'token',
        userId: 'id',
      }),
      saveStoredSession: () => {},
      patchJellyfinConfig: () => {},
      logInfo: () => {},
      logError: () => {},
      showMpvOsd: () => {},
      clearSetupWindow: () => {},
      setSetupWindow: () => {},
      encodeURIComponent,
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
  assert.equal(typeof composed.startJellyfinRemoteSession, 'function');
  assert.equal(typeof composed.stopJellyfinRemoteSession, 'function');
  assert.equal(typeof composed.runJellyfinCommand, 'function');
  assert.equal(typeof composed.openJellyfinSetupWindow, 'function');
});
