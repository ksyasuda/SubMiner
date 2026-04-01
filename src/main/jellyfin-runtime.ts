import type { CliArgs } from '../cli/args';
import {
  buildJellyfinSetupFormHtml,
  getConfiguredJellyfinSession,
  parseJellyfinSetupSubmissionUrl,
} from './runtime/domains/jellyfin';
import {
  composeJellyfinRuntimeHandlers,
  type JellyfinRuntimeComposerOptions,
} from './runtime/composers/jellyfin-runtime-composer';
import { createCreateJellyfinSetupWindowHandler } from './runtime/setup-window-factory';

// ---------------------------------------------------------------------------
// Helper: extract each dep-block's type from the composer options.
// ---------------------------------------------------------------------------

type Deps<K extends keyof JellyfinRuntimeComposerOptions> = JellyfinRuntimeComposerOptions[K];

// ---------------------------------------------------------------------------
// Resolved-config shape (extracted from composer).
// ---------------------------------------------------------------------------

type ResolvedConfigShape =
  Deps<'getResolvedJellyfinConfigMainDeps'> extends {
    getResolvedConfig: () => infer R;
  }
    ? R
    : never;
type JellyfinConfigShape = ResolvedConfigShape extends { jellyfin: infer J } ? J : never;

/** Stored-session shape (what the token store persists). */
type StoredSessionShape = { accessToken: string; userId: string };

// ---------------------------------------------------------------------------
// Public interfaces
// ---------------------------------------------------------------------------

export interface JellyfinSessionStoreLike {
  loadSession: () => StoredSessionShape | null | undefined;
  saveSession: (session: StoredSessionShape) => void;
  clearSession: () => void;
}

export interface JellyfinSetupWindowLike {
  webContents: {
    on: (event: 'will-navigate', handler: (event: unknown, url: string) => void) => void;
  };
  loadURL: (url: string) => Promise<void> | void;
  on: (event: 'closed', handler: () => void) => void;
  focus: () => void;
  close: () => void;
  isDestroyed: () => boolean;
}

/**
 * Input for createJellyfinRuntime.
 *
 * Fields whose types vary across handler files (MpvClient, Session, ClientInfo,
 * RemoteSessionService, etc.) are typed as `unknown`.  The factory body bridges
 * these to the handler-specific structural types via per-dep-block type
 * annotations (`Deps<K>`) with targeted `as` casts on the individual
 * function references.  This keeps the public-facing input surface simple and
 * avoids 7+ generic type parameters that previously required `as never` casts.
 */
export interface JellyfinRuntimeInput<
  TSetupWindow extends JellyfinSetupWindowLike = JellyfinSetupWindowLike,
> {
  getResolvedConfig: () => ResolvedConfigShape;
  getEnv: (name: string) => string | undefined;
  patchRawConfig: (patch: unknown) => void;
  defaultJellyfinConfig: JellyfinConfigShape;
  tokenStore: JellyfinSessionStoreLike;
  platform: NodeJS.Platform;
  execPath: string;
  defaultMpvLogPath: string;
  defaultMpvArgs: readonly string[];
  connectTimeoutMs: number;
  autoLaunchTimeoutMs: number;
  langPref: string;
  progressIntervalMs: number;
  ticksPerSecond: number;
  getMpvSocketPath: () => string;
  getMpvClient: () => unknown;
  setMpvClient: (client: unknown) => void;
  createMpvClient: () => unknown;
  sendMpvCommand: (client: unknown, command: Array<string | number>) => void;
  applyJellyfinMpvDefaults: (client: unknown) => void;
  showMpvOsd: (message: string) => void;
  removeSocketPath: (socketPath: string) => void;
  spawnMpv: (args: string[]) => unknown;
  wait: (delayMs: number) => Promise<void>;
  authenticateWithPassword: (
    serverUrl: string,
    username: string,
    password: string,
    clientInfo: unknown,
  ) => Promise<unknown>;
  listJellyfinLibraries: (session: unknown, clientInfo: unknown) => Promise<unknown>;
  listJellyfinItems: (session: unknown, clientInfo: unknown, params: unknown) => Promise<unknown>;
  listJellyfinSubtitleTracks: (
    session: unknown,
    clientInfo: unknown,
    itemId: string,
  ) => Promise<unknown>;
  writeJellyfinPreviewAuth: (responsePath: string, payload: unknown) => void;
  resolvePlaybackPlan: (params: unknown) => Promise<unknown>;
  convertTicksToSeconds: (ticks: number) => number;
  createRemoteSessionService: (options: unknown) => unknown;
  defaultDeviceId: string;
  defaultClientName: string;
  defaultClientVersion: string;
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TSetupWindow;
  encodeURIComponent: (value: string) => string;
  logInfo: (message: string) => void;
  logWarn: (message: string, details?: unknown) => void;
  logDebug: (message: string, details?: unknown) => void;
  logError: (message: string, error: unknown) => void;
  now?: () => number;
  schedule?: (callback: () => void, delayMs: number) => void;
}

export interface JellyfinRuntime<
  TSetupWindow extends JellyfinSetupWindowLike = JellyfinSetupWindowLike,
> {
  getResolvedJellyfinConfig: () => JellyfinConfigShape;
  reportJellyfinRemoteProgress: (forceImmediate?: boolean) => Promise<void>;
  reportJellyfinRemoteStopped: () => Promise<void>;
  startJellyfinRemoteSession: () => Promise<void>;
  stopJellyfinRemoteSession: () => Promise<void>;
  runJellyfinCommand: (args: CliArgs) => Promise<void>;
  openJellyfinSetupWindow: () => void;
  getQuitOnDisconnectArmed: () => boolean;
  clearQuitOnDisconnectArm: () => void;
  getRemoteSession: () => unknown;
  getSetupWindow: () => TSetupWindow | null;
}

export function createJellyfinRuntime<TSetupWindow extends JellyfinSetupWindowLike>(
  input: JellyfinRuntimeInput<TSetupWindow>,
): JellyfinRuntime<TSetupWindow> {
  const now = input.now ?? Date.now;
  const schedule =
    input.schedule ??
    ((callback: () => void, delayMs: number) => {
      setTimeout(callback, delayMs);
    });

  let playQuitOnDisconnectArmed = false;
  let activePlayback: unknown = null;
  let lastProgressAtMs = 0;
  let mpvAutoLaunchInFlight: Promise<boolean> | null = null;
  let remoteSession: unknown = null;
  let setupWindow: TSetupWindow | null = null;

  // Each dep block is typed with Deps<K> so TypeScript verifies structural
  // compatibility with the composer.  The `as Deps<K>[field]` casts on
  // function references bridge `unknown`-typed input methods to the
  // handler-specific structural types.  This replaces 23 `as never` casts
  // with targeted, auditable type assertions.

  const getResolvedJellyfinConfigMainDeps: Deps<'getResolvedJellyfinConfigMainDeps'> = {
    getResolvedConfig: () => input.getResolvedConfig(),
    loadStoredSession: () => input.tokenStore.loadSession(),
    getEnv: (name) => input.getEnv(name),
  };

  const getJellyfinClientInfoMainDeps: Deps<'getJellyfinClientInfoMainDeps'> = {
    getResolvedJellyfinConfig: () => input.getResolvedConfig().jellyfin,
    getDefaultJellyfinConfig: () => input.defaultJellyfinConfig,
  };

  const waitForMpvConnectedMainDeps: Deps<'waitForMpvConnectedMainDeps'> = {
    getMpvClient: input.getMpvClient as Deps<'waitForMpvConnectedMainDeps'>['getMpvClient'],
    now,
    sleep: (delayMs) => input.wait(delayMs),
  };

  const launchMpvIdleForJellyfinPlaybackMainDeps: Deps<'launchMpvIdleForJellyfinPlaybackMainDeps'> =
    {
      getSocketPath: () => input.getMpvSocketPath(),
      platform: input.platform,
      execPath: input.execPath,
      defaultMpvLogPath: input.defaultMpvLogPath,
      defaultMpvArgs: input.defaultMpvArgs,
      removeSocketPath: (socketPath) => input.removeSocketPath(socketPath),
      spawnMpv: input.spawnMpv as Deps<'launchMpvIdleForJellyfinPlaybackMainDeps'>['spawnMpv'],
      logWarn: (message, error) => input.logWarn(message, error),
      logInfo: (message) => input.logInfo(message),
    };

  const ensureMpvConnectedForJellyfinPlaybackMainDeps: Deps<'ensureMpvConnectedForJellyfinPlaybackMainDeps'> =
    {
      getMpvClient:
        input.getMpvClient as Deps<'ensureMpvConnectedForJellyfinPlaybackMainDeps'>['getMpvClient'],
      setMpvClient:
        input.setMpvClient as Deps<'ensureMpvConnectedForJellyfinPlaybackMainDeps'>['setMpvClient'],
      createMpvClient:
        input.createMpvClient as Deps<'ensureMpvConnectedForJellyfinPlaybackMainDeps'>['createMpvClient'],
      getAutoLaunchInFlight: () => mpvAutoLaunchInFlight,
      setAutoLaunchInFlight: (promise) => {
        mpvAutoLaunchInFlight = promise;
      },
      connectTimeoutMs: input.connectTimeoutMs,
      autoLaunchTimeoutMs: input.autoLaunchTimeoutMs,
    };

  const preloadJellyfinExternalSubtitlesMainDeps: Deps<'preloadJellyfinExternalSubtitlesMainDeps'> =
    {
      listJellyfinSubtitleTracks:
        input.listJellyfinSubtitleTracks as Deps<'preloadJellyfinExternalSubtitlesMainDeps'>['listJellyfinSubtitleTracks'],
      getMpvClient:
        input.getMpvClient as Deps<'preloadJellyfinExternalSubtitlesMainDeps'>['getMpvClient'],
      sendMpvCommand: (command) => input.sendMpvCommand(input.getMpvClient(), command),
      wait: (delayMs) => input.wait(delayMs),
      logDebug: (message, error) => input.logDebug(message, error),
    };

  const playJellyfinItemInMpvMainDeps: Deps<'playJellyfinItemInMpvMainDeps'> = {
    getMpvClient: input.getMpvClient as Deps<'playJellyfinItemInMpvMainDeps'>['getMpvClient'],
    resolvePlaybackPlan:
      input.resolvePlaybackPlan as Deps<'playJellyfinItemInMpvMainDeps'>['resolvePlaybackPlan'],
    applyJellyfinMpvDefaults:
      input.applyJellyfinMpvDefaults as Deps<'playJellyfinItemInMpvMainDeps'>['applyJellyfinMpvDefaults'],
    sendMpvCommand: (command) => input.sendMpvCommand(input.getMpvClient(), command),
    armQuitOnDisconnect: () => {
      playQuitOnDisconnectArmed = false;
      schedule(() => {
        playQuitOnDisconnectArmed = true;
      }, 3000);
    },
    schedule: (callback, delayMs) => {
      schedule(callback, delayMs);
    },
    convertTicksToSeconds: (ticks) => input.convertTicksToSeconds(ticks),
    setActivePlayback: (state) => {
      activePlayback = state;
    },
    setLastProgressAtMs: (value) => {
      lastProgressAtMs = value;
    },
    reportPlaying: (payload) => {
      const session = remoteSession as { reportPlaying?: (payload: unknown) => unknown } | null;
      if (typeof session?.reportPlaying === 'function') {
        void session.reportPlaying(payload);
      }
    },
    showMpvOsd: (message) => input.showMpvOsd(message),
  };

  const remoteComposerBase: Omit<Deps<'remoteComposerOptions'>, 'getConfiguredSession'> = {
    logWarn: (message) => input.logWarn(message),
    getMpvClient: input.getMpvClient as Deps<'remoteComposerOptions'>['getMpvClient'],
    sendMpvCommand: input.sendMpvCommand as Deps<'remoteComposerOptions'>['sendMpvCommand'],
    jellyfinTicksToSeconds: (ticks) => input.convertTicksToSeconds(ticks),
    getActivePlayback: () =>
      activePlayback as ReturnType<Deps<'remoteComposerOptions'>['getActivePlayback']>,
    clearActivePlayback: () => {
      activePlayback = null;
    },
    getSession: () => remoteSession as ReturnType<Deps<'remoteComposerOptions'>['getSession']>,
    getNow: now,
    getLastProgressAtMs: () => lastProgressAtMs,
    setLastProgressAtMs: (value) => {
      lastProgressAtMs = value;
    },
    progressIntervalMs: input.progressIntervalMs,
    ticksPerSecond: input.ticksPerSecond,
    logDebug: (message, error) => input.logDebug(message, error),
  };

  const handleJellyfinAuthCommandsMainDeps: Deps<'handleJellyfinAuthCommandsMainDeps'> = {
    patchRawConfig: (patch) => input.patchRawConfig(patch),
    authenticateWithPassword:
      input.authenticateWithPassword as Deps<'handleJellyfinAuthCommandsMainDeps'>['authenticateWithPassword'],
    saveStoredSession: (session) => input.tokenStore.saveSession(session as StoredSessionShape),
    clearStoredSession: () => input.tokenStore.clearSession(),
    logInfo: (message) => input.logInfo(message),
  };

  const handleJellyfinListCommandsMainDeps: Deps<'handleJellyfinListCommandsMainDeps'> = {
    listJellyfinLibraries:
      input.listJellyfinLibraries as Deps<'handleJellyfinListCommandsMainDeps'>['listJellyfinLibraries'],
    listJellyfinItems:
      input.listJellyfinItems as Deps<'handleJellyfinListCommandsMainDeps'>['listJellyfinItems'],
    listJellyfinSubtitleTracks:
      input.listJellyfinSubtitleTracks as Deps<'handleJellyfinListCommandsMainDeps'>['listJellyfinSubtitleTracks'],
    writeJellyfinPreviewAuth: (responsePath, payload) =>
      input.writeJellyfinPreviewAuth(responsePath, payload),
    logInfo: (message) => input.logInfo(message),
  };

  const handleJellyfinPlayCommandMainDeps: Deps<'handleJellyfinPlayCommandMainDeps'> = {
    logWarn: (message) => input.logWarn(message),
  };

  const handleJellyfinRemoteAnnounceCommandMainDeps: Deps<'handleJellyfinRemoteAnnounceCommandMainDeps'> =
    {
      getRemoteSession: () =>
        remoteSession as ReturnType<
          Deps<'handleJellyfinRemoteAnnounceCommandMainDeps'>['getRemoteSession']
        >,
      logInfo: (message) => input.logInfo(message),
      logWarn: (message) => input.logWarn(message),
    };

  const startJellyfinRemoteSessionMainDeps: Deps<'startJellyfinRemoteSessionMainDeps'> = {
    getCurrentSession: () =>
      remoteSession as ReturnType<Deps<'startJellyfinRemoteSessionMainDeps'>['getCurrentSession']>,
    setCurrentSession: (session) => {
      remoteSession = session;
    },
    createRemoteSessionService:
      input.createRemoteSessionService as Deps<'startJellyfinRemoteSessionMainDeps'>['createRemoteSessionService'],
    defaultDeviceId: input.defaultDeviceId,
    defaultClientName: input.defaultClientName,
    defaultClientVersion: input.defaultClientVersion,
    logInfo: (message) => input.logInfo(message),
    logWarn: (message, details) => input.logWarn(message, details),
  };

  const stopJellyfinRemoteSessionMainDeps: Deps<'stopJellyfinRemoteSessionMainDeps'> = {
    getCurrentSession: () =>
      remoteSession as ReturnType<Deps<'stopJellyfinRemoteSessionMainDeps'>['getCurrentSession']>,
    setCurrentSession: (session) => {
      remoteSession = session;
    },
    clearActivePlayback: () => {
      activePlayback = null;
    },
  };

  const runJellyfinCommandMainDeps: Deps<'runJellyfinCommandMainDeps'> = {
    defaultServerUrl: input.defaultJellyfinConfig.serverUrl,
  };

  const maybeFocusExistingJellyfinSetupWindowMainDeps: Deps<'maybeFocusExistingJellyfinSetupWindowMainDeps'> =
    {
      getSetupWindow: () => setupWindow,
    };

  const openJellyfinSetupWindowMainDeps: Deps<'openJellyfinSetupWindowMainDeps'> = {
    createSetupWindow: createCreateJellyfinSetupWindowHandler({
      createBrowserWindow: (options) => input.createBrowserWindow(options),
    }),
    buildSetupFormHtml: (defaultServer, defaultUser) =>
      buildJellyfinSetupFormHtml(defaultServer, defaultUser),
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword:
      input.authenticateWithPassword as Deps<'openJellyfinSetupWindowMainDeps'>['authenticateWithPassword'],
    saveStoredSession: (session) => input.tokenStore.saveSession(session as StoredSessionShape),
    patchJellyfinConfig: (session) => {
      const jellyfinSession = session as { serverUrl?: string; username?: string };
      input.patchRawConfig({
        jellyfin: {
          enabled: true,
          serverUrl: jellyfinSession.serverUrl,
          username: jellyfinSession.username,
        },
      });
    },
    logInfo: (message) => input.logInfo(message),
    logError: (message, error) => input.logError(message, error),
    showMpvOsd: (message) => input.showMpvOsd(message),
    clearSetupWindow: () => {
      setupWindow = null;
    },
    setSetupWindow: (window) => {
      setupWindow = window as TSetupWindow | null;
    },
    encodeURIComponent: (value) => input.encodeURIComponent(value),
  };

  const runtime = composeJellyfinRuntimeHandlers({
    getResolvedJellyfinConfigMainDeps,
    getJellyfinClientInfoMainDeps,
    waitForMpvConnectedMainDeps,
    launchMpvIdleForJellyfinPlaybackMainDeps,
    ensureMpvConnectedForJellyfinPlaybackMainDeps,
    preloadJellyfinExternalSubtitlesMainDeps,
    playJellyfinItemInMpvMainDeps,
    remoteComposerOptions: {
      ...remoteComposerBase,
      getConfiguredSession: () => getConfiguredJellyfinSession(runtime.getResolvedJellyfinConfig()),
    },
    handleJellyfinAuthCommandsMainDeps,
    handleJellyfinListCommandsMainDeps,
    handleJellyfinPlayCommandMainDeps,
    handleJellyfinRemoteAnnounceCommandMainDeps,
    startJellyfinRemoteSessionMainDeps,
    stopJellyfinRemoteSessionMainDeps,
    runJellyfinCommandMainDeps,
    maybeFocusExistingJellyfinSetupWindowMainDeps,
    openJellyfinSetupWindowMainDeps,
  });

  return {
    getResolvedJellyfinConfig: () => runtime.getResolvedJellyfinConfig(),
    reportJellyfinRemoteProgress: async (forceImmediate) => {
      await runtime.reportJellyfinRemoteProgress(forceImmediate);
    },
    reportJellyfinRemoteStopped: async () => {
      await runtime.reportJellyfinRemoteStopped();
    },
    startJellyfinRemoteSession: async () => {
      await runtime.startJellyfinRemoteSession();
    },
    stopJellyfinRemoteSession: async () => {
      await runtime.stopJellyfinRemoteSession();
    },
    runJellyfinCommand: async (args) => {
      await runtime.runJellyfinCommand(args);
    },
    openJellyfinSetupWindow: () => {
      runtime.openJellyfinSetupWindow();
    },
    getQuitOnDisconnectArmed: () => playQuitOnDisconnectArmed,
    clearQuitOnDisconnectArm: () => {
      playQuitOnDisconnectArmed = false;
    },
    getRemoteSession: () => remoteSession,
    getSetupWindow: () => setupWindow,
  };
}
