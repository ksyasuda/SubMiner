import fs from 'node:fs';
import os from 'node:os';
import type { BrowserWindow, BrowserWindowConstructorOptions, IpcMain } from 'electron';
import { DEFAULT_CONFIG } from '../../config/definitions';
import type { ConfigService } from '../../config/service';
import {
  authenticateWithPasswordRuntime,
  JellyfinPlaybackReporter,
  JellyfinRemoteSessionService,
  jellyfinTicksToSecondsRuntime,
  listJellyfinItemsRuntime,
  listJellyfinLibrariesRuntime,
  listJellyfinSubtitleTracksRuntime,
  resolveJellyfinPlaybackPlanRuntime,
  sendMpvCommandRuntime,
  type MpvIpcClient,
} from '../../core/services';
import type { JellyfinTokenStore } from '../../core/services/jellyfin-token-store';
import type { createLogger } from '../../logger';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { writeJellyfinCliResponse } from '../../shared/jellyfin-cli-response';
import { isLogFileEnabled } from '../../shared/log-files';
import { composeJellyfinRuntimeHandlers } from '../runtime/composers';
import {
  buildJellyfinSetupFormHtml,
  createApplyJellyfinMpvDefaultsHandler,
  getConfiguredJellyfinSession,
  mergeJellyfinRecentServers,
  parseJellyfinSetupSubmissionUrl,
  persistJellyfinAuthSession,
  type ActiveJellyfinRemotePlaybackState,
} from '../runtime/domains/jellyfin';
import { detectInstalledMpvPlugin } from '../runtime/first-run-setup-plugin';
import {
  DEFAULT_JELLYFIN_CLIENT_NAME,
  DEFAULT_JELLYFIN_CLIENT_VERSION,
  createHostDerivedJellyfinDeviceId,
  resolveJellyfinRemoteDeviceName,
} from '../runtime/jellyfin-device-identity';
import { createJellyfinSubtitleCacheIo } from '../runtime/jellyfin-subtitle-cache-io';
import { clearJellyfinAuthSessionAndRefreshTray as clearJellyfinAuthSessionAndRefreshTrayRuntime } from '../runtime/jellyfin-tray-discovery';
import { resolveMpvExecutablePath, spawnMpvProcess } from '../runtime/mpv-process';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createCreateJellyfinSetupWindowHandler } from '../runtime/setup-window-factory';
import type { AppState } from '../state';

const JELLYFIN_LANG_PREF = 'ja,jp,jpn,japanese,en,eng,english,enUS,en-US';
const JELLYFIN_TICKS_PER_SECOND = 10_000_000;
const JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS = 3000;
const JELLYFIN_REMOTE_STARTUP_STOP_GRACE_MS = 10_000;
const JELLYFIN_MPV_CONNECT_TIMEOUT_MS = 3000;
const JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS = 20000;
const MPV_JELLYFIN_DEFAULT_ARGS = [
  '--sub-auto=no',
  '--sub-file-paths=.;subs;subtitles',
  '--sid=no',
  '--secondary-sid=no',
  '--sub-visibility=no',
  '--secondary-sub-visibility=no',
  '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
  '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
] as const;

type JellyfinComposerOptions = Parameters<typeof composeJellyfinRuntimeHandlers>[0];
type LaunchMpvIdleDeps = JellyfinComposerOptions['launchMpvIdleForJellyfinPlaybackMainDeps'];

/**
 * What the Jellyfin domain needs from the rest of the app. Siblings built after this assembly in
 * `main.ts` are reached through getters or callbacks, which are only called once the app runs.
 */
export interface JellyfinAssemblyDeps {
  appState: Pick<
    AppState,
    | 'mpvClient'
    | 'mpvSocketPath'
    | 'immersionTracker'
    | 'jellyfinRemoteSession'
    | 'jellyfinSetupWindow'
  >;
  configService: Pick<ConfigService, 'getConfig' | 'patchRawConfig'>;
  jellyfinTokenStore: JellyfinTokenStore;
  logger: Pick<ReturnType<typeof createLogger>, 'info' | 'warn' | 'debug' | 'error'>;
  defaultMpvLogPath: string;
  getAppDataPath: () => string;
  getRuntimePluginEntrypoint: NonNullable<LaunchMpvIdleDeps['getRuntimePluginEntrypoint']>;
  getMpvPluginRuntimeConfig: NonNullable<LaunchMpvIdleDeps['getPluginRuntimeConfig']>;
  createMpvClient: () => MpvIpcClient;
  initSubtitlePrefetch: NonNullable<
    JellyfinComposerOptions['preloadJellyfinExternalSubtitlesMainDeps']['initSubtitlePrefetch']
  >;
  showVisibleOverlay: () => void;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'showConfiguredStatusNotification'
  >;
  updateCurrentMediaTitle: (title: string) => void;
  ensureImmersionTrackerStarted: () => void;
  refreshTrayMenu: () => void;
  setupPreloadPath: string;
  createBrowserWindow: (options: BrowserWindowConstructorOptions) => BrowserWindow;
  ipcMain: Pick<IpcMain, 'handle' | 'removeHandler'>;
}

export type JellyfinAssembly = ReturnType<typeof createJellyfinAssembly>;

/**
 * Owns Jellyfin playback, remote-control session, CLI commands, setup window and tray discovery,
 * plus the playback state they share. `main.ts` builds it once and calls through the result.
 */
export function createJellyfinAssembly(deps: JellyfinAssemblyDeps) {
  const { appState, configService, jellyfinTokenStore, logger } = deps;

  const jellyfinSubtitleCacheIo = createJellyfinSubtitleCacheIo({
    tmpDir: () => os.tmpdir(),
    makeTempDir: (prefix) => fs.promises.mkdtemp(prefix),
    writeFile: (filePath, bytes) => fs.promises.writeFile(filePath, bytes),
    removeDir: (dir, options) => {
      fs.rmSync(dir, options);
    },
    fetch: (url) => fetch(url),
  });

  let jellyfinPlayQuitOnDisconnectArmed = false;
  let activeJellyfinRemotePlayback: ActiveJellyfinRemotePlaybackState | null = null;
  let jellyfinRemoteLastProgressAtMs = 0;
  let jellyfinMpvAutoLaunchInFlight: Promise<boolean> | null = null;

  const applyJellyfinMpvDefaultsHandler = createApplyJellyfinMpvDefaultsHandler({
    sendMpvCommandRuntime: (client, command) => sendMpvCommandRuntime(client, command),
    jellyfinLangPref: JELLYFIN_LANG_PREF,
  });

  const {
    getResolvedJellyfinConfig,
    reportJellyfinRemoteProgress,
    reportJellyfinRemoteStopped,
    startJellyfinRemoteSession,
    stopJellyfinRemoteSession,
    cleanupJellyfinSubtitleCache,
    runJellyfinCommand,
    openJellyfinSetupWindow,
    ensureMpvConnectedForPlayback,
  } = composeJellyfinRuntimeHandlers({
    getResolvedJellyfinConfigMainDeps: {
      getResolvedConfig: () => configService.getConfig(),
      loadStoredSession: () => jellyfinTokenStore.loadSession(),
      getEnv: (name) => process.env[name],
    },
    getJellyfinClientInfoMainDeps: {
      getResolvedJellyfinConfig: () => getResolvedJellyfinConfig(),
      getHostName: () => os.hostname(),
      defaultClientName: DEFAULT_JELLYFIN_CLIENT_NAME,
      defaultClientVersion: DEFAULT_JELLYFIN_CLIENT_VERSION,
    },
    waitForMpvConnectedMainDeps: {
      getMpvClient: () => appState.mpvClient,
      now: () => Date.now(),
      sleep: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
    },
    launchMpvIdleForJellyfinPlaybackMainDeps: {
      getMpvExecutablePath: () =>
        resolveMpvExecutablePath(configService.getConfig().mpv.executablePath),
      getSocketPath: () => appState.mpvSocketPath,
      getLaunchMode: () => configService.getConfig().mpv.launchMode,
      platform: process.platform,
      execPath: process.execPath,
      getRuntimePluginEntrypoint: () => deps.getRuntimePluginEntrypoint(),
      getInstalledPluginDetection: (mpvExecutablePath) =>
        detectInstalledMpvPlugin({
          platform: process.platform,
          homeDir: os.homedir(),
          xdgConfigHome: process.env.XDG_CONFIG_HOME,
          appDataDir: deps.getAppDataPath(),
          mpvExecutablePath,
        }),
      getPluginRuntimeConfig: () => deps.getMpvPluginRuntimeConfig(),
      getDefaultMpvLogPath: () => (isLogFileEnabled('mpv') ? deps.defaultMpvLogPath : ''),
      defaultMpvArgs: MPV_JELLYFIN_DEFAULT_ARGS,
      removeSocketPath: (socketPath) => {
        fs.rmSync(socketPath, { force: true });
      },
      spawnMpv: spawnMpvProcess,
      logWarn: (message, error) => logger.warn(message, error),
      logInfo: (message) => logger.info(message),
    },
    ensureMpvConnectedForJellyfinPlaybackMainDeps: {
      getMpvClient: () => appState.mpvClient,
      setMpvClient: (client) => {
        appState.mpvClient = client as MpvIpcClient | null;
      },
      createMpvClient: () => deps.createMpvClient(),
      getAutoLaunchInFlight: () => jellyfinMpvAutoLaunchInFlight,
      setAutoLaunchInFlight: (promise) => {
        jellyfinMpvAutoLaunchInFlight = promise;
      },
      connectTimeoutMs: JELLYFIN_MPV_CONNECT_TIMEOUT_MS,
      autoLaunchTimeoutMs: JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS,
    },
    preloadJellyfinExternalSubtitlesMainDeps: {
      listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
        listJellyfinSubtitleTracksRuntime(session, clientInfo, itemId),
      getMpvClient: () => appState.mpvClient,
      sendMpvCommand: (command) => {
        sendMpvCommandRuntime(appState.mpvClient, command);
      },
      wait: (ms) => new Promise<void>((resolve) => setTimeout(resolve, ms)),
      cacheSubtitleTrack: (track) => jellyfinSubtitleCacheIo.cacheSubtitleTrack(track),
      cleanupCachedSubtitles: (dirs) => jellyfinSubtitleCacheIo.cleanupCachedSubtitles(dirs),
      initSubtitlePrefetch: (sourcePath) => deps.initSubtitlePrefetch(sourcePath),
      logDebug: (message, error) => {
        logger.debug(message, error);
      },
    },
    playJellyfinItemInMpvMainDeps: {
      getMpvClient: () => appState.mpvClient,
      resolvePlaybackPlan: (params) =>
        resolveJellyfinPlaybackPlanRuntime(
          params.session,
          params.clientInfo,
          params.jellyfinConfig as ReturnType<typeof getResolvedJellyfinConfig>,
          {
            itemId: params.itemId,
            audioStreamIndex: params.audioStreamIndex ?? undefined,
            subtitleStreamIndex: params.subtitleStreamIndex ?? undefined,
          },
        ),
      applyJellyfinMpvDefaults: (mpvClient) => applyJellyfinMpvDefaultsHandler(mpvClient),
      showVisibleOverlay: () => deps.showVisibleOverlay(),
      sendMpvCommand: (command) => sendMpvCommandRuntime(appState.mpvClient, command),
      armQuitOnDisconnect: () => {
        jellyfinPlayQuitOnDisconnectArmed = false;
        setTimeout(() => {
          jellyfinPlayQuitOnDisconnectArmed = true;
        }, 3000);
      },
      schedule: (callback, delayMs) => {
        setTimeout(callback, delayMs);
      },
      convertTicksToSeconds: (ticks) => jellyfinTicksToSecondsRuntime(ticks),
      setActivePlayback: (state) => {
        activeJellyfinRemotePlayback = {
          ...(state as ActiveJellyfinRemotePlaybackState),
          stopReportsAfterMs:
            state.stopReportsAfterMs ?? Date.now() + JELLYFIN_REMOTE_STARTUP_STOP_GRACE_MS,
        };
      },
      setLastProgressAtMs: (value) => {
        jellyfinRemoteLastProgressAtMs = value;
      },
      // Same device identity as the cast session, so Jellyfin attributes reports to one session.
      createPlaybackReporter: ({ session, clientInfo }) =>
        new JellyfinPlaybackReporter({
          serverUrl: session.serverUrl,
          accessToken: session.accessToken,
          deviceId: clientInfo.deviceId,
          clientName: clientInfo.clientName,
          clientVersion: clientInfo.clientVersion,
          deviceName: resolveJellyfinRemoteDeviceName({ hostName: os.hostname() }),
          logWarn: (message, details) => logger.warn(message, details),
        }),
      showMpvOsd: (text) => {
        deps
          .getOverlayNotifications()
          .showConfiguredStatusNotification(text, { title: 'Jellyfin' });
      },
      updateCurrentMediaTitle: (title) => {
        deps.updateCurrentMediaTitle(title);
      },
      recordJellyfinPlaybackMetadata: (metadata) => {
        deps.ensureImmersionTrackerStarted();
        appState.immersionTracker?.recordJellyfinPlaybackMetadata(metadata);
      },
    },
    remoteComposerOptions: {
      getConfiguredSession: () => getConfiguredJellyfinSession(getResolvedJellyfinConfig()),
      logWarn: (message) => logger.warn(message),
      getMpvClient: () => appState.mpvClient,
      sendMpvCommand: (client, command) => sendMpvCommandRuntime(client as MpvIpcClient, command),
      jellyfinTicksToSeconds: (ticks) => jellyfinTicksToSecondsRuntime(ticks),
      getActivePlayback: () => activeJellyfinRemotePlayback,
      clearActivePlayback: () => {
        activeJellyfinRemotePlayback = null;
      },
      getNow: () => Date.now(),
      getLastProgressAtMs: () => jellyfinRemoteLastProgressAtMs,
      setLastProgressAtMs: (value) => {
        jellyfinRemoteLastProgressAtMs = value;
      },
      progressIntervalMs: JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS,
      ticksPerSecond: JELLYFIN_TICKS_PER_SECOND,
      logDebug: (message, error) => logger.debug(message, error),
    },
    handleJellyfinAuthCommandsMainDeps: {
      patchRawConfig: (patch) => {
        configService.patchRawConfig(patch);
        deps.refreshTrayMenu();
      },
      authenticateWithPassword: (serverUrl, username, password, clientInfo) =>
        authenticateWithPasswordRuntime(serverUrl, username, password, clientInfo),
      saveStoredSession: (session) => jellyfinTokenStore.saveSession(session),
      clearStoredSession: () =>
        clearJellyfinAuthSessionAndRefreshTrayRuntime(getJellyfinTrayDiscoveryDeps()),
      logInfo: (message) => logger.info(message),
    },
    handleJellyfinListCommandsMainDeps: {
      listJellyfinLibraries: (session, clientInfo) =>
        listJellyfinLibrariesRuntime(session, clientInfo),
      listJellyfinItems: (session, clientInfo, params) =>
        listJellyfinItemsRuntime(session, clientInfo, params),
      listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
        listJellyfinSubtitleTracksRuntime(session, clientInfo, itemId),
      writeJellyfinResponse: writeJellyfinCliResponse,
      logInfo: (message) => logger.info(message),
    },
    handleJellyfinPlayCommandMainDeps: {
      logWarn: (message) => logger.warn(message),
    },
    handleJellyfinRemoteAnnounceCommandMainDeps: {
      getRemoteSession: () => appState.jellyfinRemoteSession,
      logInfo: (message) => logger.info(message),
      logWarn: (message) => logger.warn(message),
    },
    startJellyfinRemoteSessionMainDeps: {
      getCurrentSession: () => appState.jellyfinRemoteSession,
      setCurrentSession: (session) => {
        appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
      },
      createRemoteSessionService: (options) => new JellyfinRemoteSessionService(options),
      getHostName: () => os.hostname(),
      defaultDeviceId: createHostDerivedJellyfinDeviceId(os.hostname()),
      defaultClientName: DEFAULT_JELLYFIN_CLIENT_NAME,
      defaultClientVersion: DEFAULT_JELLYFIN_CLIENT_VERSION,
      logInfo: (message) => logger.info(message),
      logWarn: (message, details) => logger.warn(message, details),
      onSessionStateChanged: () => deps.refreshTrayMenu(),
    },
    stopJellyfinRemoteSessionMainDeps: {
      getCurrentSession: () => appState.jellyfinRemoteSession,
      setCurrentSession: (session) => {
        appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
      },
      clearActivePlayback: () => {
        activeJellyfinRemotePlayback = null;
      },
      onSessionStateChanged: () => deps.refreshTrayMenu(),
    },
    runJellyfinCommandMainDeps: {
      defaultServerUrl: DEFAULT_CONFIG.jellyfin.serverUrl,
      writeJellyfinResponse: writeJellyfinCliResponse,
    },
    maybeFocusExistingJellyfinSetupWindowMainDeps: {
      getSetupWindow: () => appState.jellyfinSetupWindow,
    },
    openJellyfinSetupWindowMainDeps: {
      createSetupWindow: createCreateJellyfinSetupWindowHandler({
        createBrowserWindow: (options) => deps.createBrowserWindow(options),
        preloadPath: deps.setupPreloadPath,
      }),
      buildSetupFormHtml: (state) => buildJellyfinSetupFormHtml(state),
      parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
      authenticateWithPassword: (server, username, password, clientInfo) =>
        authenticateWithPasswordRuntime(server, username, password, clientInfo),
      saveStoredSession: (session) => jellyfinTokenStore.saveSession(session),
      clearStoredSession: () =>
        clearJellyfinAuthSessionAndRefreshTrayRuntime(getJellyfinTrayDiscoveryDeps()),
      patchJellyfinConfig: (session) => {
        const recentServers = mergeJellyfinRecentServers(
          session.serverUrl,
          configService.getConfig().jellyfin.recentServers || [],
        );
        configService.patchRawConfig({
          jellyfin: {
            enabled: true,
            serverUrl: session.serverUrl,
            username: session.username,
            recentServers,
          },
        });
        deps.refreshTrayMenu();
      },
      persistAuthenticatedSession: (session, clientInfo) =>
        persistJellyfinAuthSession({
          session,
          clientInfo,
          existingRecentServers: configService.getConfig().jellyfin.recentServers || [],
          saveStoredSession: (storedSession) => jellyfinTokenStore.saveSession(storedSession),
          patchRawConfig: (patch) => {
            configService.patchRawConfig(patch);
            deps.refreshTrayMenu();
          },
        }),
      logInfo: (message) => logger.info(message),
      logError: (message, error) => logger.error(message, error),
      showMpvOsd: (message) =>
        deps
          .getOverlayNotifications()
          .showConfiguredStatusNotification(message, { title: 'Jellyfin' }),
      clearSetupWindow: () => {
        appState.jellyfinSetupWindow = null;
      },
      setSetupWindow: (window) => {
        appState.jellyfinSetupWindow = window as BrowserWindow;
      },
      registerSetupIpcHandler: (handler) => {
        const channel = IPC_CHANNELS.request.jellyfinSetupSubmit;
        deps.ipcMain.removeHandler(channel);
        deps.ipcMain.handle(channel, async (event, payload) => {
          const setupWindow = appState.jellyfinSetupWindow;
          if (!setupWindow || event.sender !== setupWindow.webContents) {
            return {
              handled: false,
              statusMessage: 'This Jellyfin setup window is no longer active.',
              statusKind: 'error',
            };
          }
          return handler(payload);
        });
        return () => {
          deps.ipcMain.removeHandler(channel);
        };
      },
      encodeURIComponent: (value) => encodeURIComponent(value),
      defaultServerUrl: DEFAULT_CONFIG.jellyfin.serverUrl || 'http://127.0.0.1:8096',
      hasStoredSession: () => Boolean(jellyfinTokenStore.loadSession()),
    },
  });

  function getJellyfinTrayDiscoveryDeps() {
    return {
      getResolvedJellyfinConfig: () => getResolvedJellyfinConfig(),
      getRemoteSession: () => appState.jellyfinRemoteSession,
      clearStoredSession: () => jellyfinTokenStore.clearSession(),
      stopRemoteSession: () => stopJellyfinRemoteSession(),
      startRemoteSession: (options: { explicit: true }) => startJellyfinRemoteSession(options),
      refreshTrayMenu: () => deps.refreshTrayMenu(),
      logger,
      showMpvOsd: (message: string) =>
        deps
          .getOverlayNotifications()
          .showConfiguredStatusNotification(message, { title: 'Jellyfin' }),
    };
  }

  return {
    reportJellyfinRemoteProgress,
    reportJellyfinRemoteStopped,
    startJellyfinRemoteSession,
    stopJellyfinRemoteSession,
    cleanupJellyfinSubtitleCache,
    runJellyfinCommand,
    openJellyfinSetupWindow,
    ensureMpvConnectedForPlayback,
    getTrayDiscoveryDeps: getJellyfinTrayDiscoveryDeps,
    getActivePlayback: () => activeJellyfinRemotePlayback,
    getQuitOnDisconnectArmed: () => jellyfinPlayQuitOnDisconnectArmed,
  };
}
