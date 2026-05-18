import type { BrowserWindow } from 'electron';
import { ConfigStartupParseError } from '../../config';

export interface AppLifecycleShape {
  requestSingleInstanceLock: () => boolean;
  quit: () => void;
  exit: (code?: number) => void;
  on: (event: string, listener: (...args: unknown[]) => void) => unknown;
  whenReady: () => Promise<void>;
}

export interface OverlayModalInputStateShape {
  getModalInputExclusive: () => boolean;
  handleModalInputStateChange: (isActive: boolean) => void;
}

export interface MainBootServicesParams<
  TConfigService,
  TAnilistTokenStore,
  TJellyfinTokenStore,
  TAnilistUpdateQueue,
  TSubtitleWebSocket,
  TLogger,
  TRuntimeRegistry,
  TOverlayManager,
  TOverlayModalInputState,
  TOverlayContentMeasurementStore,
  TOverlayModalRuntime,
  TAppState,
  TAppLifecycleApp,
> {
  platform: NodeJS.Platform;
  argv: string[];
  appDataDir: string | undefined;
  xdgConfigHome: string | undefined;
  homeDir: string;
  defaultMpvLogFile: string;
  envMpvLog: string | undefined;
  defaultTexthookerPort: number;
  getDefaultSocketPath: () => string;
  resolveConfigDir: (input: {
    platform: NodeJS.Platform;
    appDataDir: string | undefined;
    xdgConfigHome: string | undefined;
    homeDir: string;
    existsSync: (targetPath: string) => boolean;
  }) => string;
  existsSync: (targetPath: string) => boolean;
  mkdirSync: (targetPath: string, options: { recursive: true }) => void;
  joinPath: (...parts: string[]) => string;
  app: {
    setPath: (name: string, value: string) => void;
    quit: () => void;
    exit: (code?: number) => void;
    // eslint-disable-next-line @typescript-eslint/no-unsafe-function-type -- Electron App.on has 50+ overloaded signatures
    on: Function;
    whenReady: () => Promise<void>;
  };
  shouldBypassSingleInstanceLock: () => boolean;
  requestSingleInstanceLockEarly: () => boolean;
  registerSecondInstanceHandlerEarly: (listener: (_event: unknown, argv: string[]) => void) => void;
  onConfigStartupParseError: (error: ConfigStartupParseError) => void;
  createConfigService: (configDir: string) => TConfigService;
  createAnilistTokenStore: (targetPath: string) => TAnilistTokenStore;
  createJellyfinTokenStore: (targetPath: string) => TJellyfinTokenStore;
  createAnilistUpdateQueue: (targetPath: string) => TAnilistUpdateQueue;
  createSubtitleWebSocket: () => TSubtitleWebSocket;
  createLogger: (scope: string) => TLogger & {
    warn: (message: string) => void;
    info: (message: string) => void;
    error: (message: string, details?: unknown) => void;
  };
  createMainRuntimeRegistry: () => TRuntimeRegistry;
  createOverlayManager: () => TOverlayManager;
  createOverlayModalInputState: (params: {
    getModalWindow: () => BrowserWindow | null;
    syncOverlayShortcutsForModal: (isActive: boolean) => void;
    syncOverlayVisibilityForModal: () => void;
    restoreMainWindowFocus?: () => void;
  }) => TOverlayModalInputState;
  createOverlayContentMeasurementStore: (params: {
    logger: TLogger;
  }) => TOverlayContentMeasurementStore;
  getSyncOverlayShortcutsForModal: () => (isActive: boolean) => void;
  getSyncOverlayVisibilityForModal: () => () => void;
  createOverlayModalRuntime: (params: {
    overlayManager: TOverlayManager;
    overlayModalInputState: TOverlayModalInputState;
    onModalStateChange: (isActive: boolean) => void;
  }) => TOverlayModalRuntime;
  createAppState: (input: { mpvSocketPath: string; texthookerPort: number }) => TAppState;
}

export interface MainBootServicesResult<
  TConfigService,
  TAnilistTokenStore,
  TJellyfinTokenStore,
  TAnilistUpdateQueue,
  TSubtitleWebSocket,
  TLogger,
  TRuntimeRegistry,
  TOverlayManager,
  TOverlayModalInputState,
  TOverlayContentMeasurementStore,
  TOverlayModalRuntime,
  TAppState,
  TAppLifecycleApp,
> {
  configDir: string;
  userDataPath: string;
  defaultMpvLogPath: string;
  defaultImmersionDbPath: string;
  configService: TConfigService;
  anilistTokenStore: TAnilistTokenStore;
  jellyfinTokenStore: TJellyfinTokenStore;
  anilistUpdateQueue: TAnilistUpdateQueue;
  subtitleWsService: TSubtitleWebSocket;
  annotationSubtitleWsService: TSubtitleWebSocket;
  logger: TLogger;
  runtimeRegistry: TRuntimeRegistry;
  overlayManager: TOverlayManager;
  overlayModalInputState: TOverlayModalInputState;
  overlayContentMeasurementStore: TOverlayContentMeasurementStore;
  overlayModalRuntime: TOverlayModalRuntime;
  appState: TAppState;
  appLifecycleApp: TAppLifecycleApp;
}

export function createMainBootServices<
  TConfigService,
  TAnilistTokenStore,
  TJellyfinTokenStore,
  TAnilistUpdateQueue,
  TSubtitleWebSocket,
  TLogger,
  TRuntimeRegistry,
  TOverlayManager extends {
    getMainWindow: () => BrowserWindow | null;
    getModalWindow: () => BrowserWindow | null;
  },
  TOverlayModalInputState extends OverlayModalInputStateShape,
  TOverlayContentMeasurementStore,
  TOverlayModalRuntime,
  TAppState,
  TAppLifecycleApp extends AppLifecycleShape,
>(
  params: MainBootServicesParams<
    TConfigService,
    TAnilistTokenStore,
    TJellyfinTokenStore,
    TAnilistUpdateQueue,
    TSubtitleWebSocket,
    TLogger,
    TRuntimeRegistry,
    TOverlayManager,
    TOverlayModalInputState,
    TOverlayContentMeasurementStore,
    TOverlayModalRuntime,
    TAppState,
    TAppLifecycleApp
  >,
): MainBootServicesResult<
  TConfigService,
  TAnilistTokenStore,
  TJellyfinTokenStore,
  TAnilistUpdateQueue,
  TSubtitleWebSocket,
  TLogger,
  TRuntimeRegistry,
  TOverlayManager,
  TOverlayModalInputState,
  TOverlayContentMeasurementStore,
  TOverlayModalRuntime,
  TAppState,
  TAppLifecycleApp
> {
  const configDir = params.resolveConfigDir({
    platform: params.platform,
    appDataDir: params.appDataDir,
    xdgConfigHome: params.xdgConfigHome,
    homeDir: params.homeDir,
    existsSync: params.existsSync,
  });
  const userDataPath = configDir;
  const defaultMpvLogPath = params.envMpvLog?.trim() || params.defaultMpvLogFile;
  const defaultImmersionDbPath = params.joinPath(userDataPath, 'immersion.sqlite');

  const configService = (() => {
    try {
      return params.createConfigService(configDir);
    } catch (error) {
      if (error instanceof ConfigStartupParseError) {
        params.onConfigStartupParseError(error);
      }
      throw error;
    }
  })();

  const anilistTokenStore = params.createAnilistTokenStore(
    params.joinPath(userDataPath, 'anilist-token-store.json'),
  );
  const jellyfinTokenStore = params.createJellyfinTokenStore(
    params.joinPath(userDataPath, 'jellyfin-token-store.json'),
  );
  const anilistUpdateQueue = params.createAnilistUpdateQueue(
    params.joinPath(userDataPath, 'anilist-retry-queue.json'),
  );
  const subtitleWsService = params.createSubtitleWebSocket();
  const annotationSubtitleWsService = params.createSubtitleWebSocket();
  const logger = params.createLogger('main');
  const runtimeRegistry = params.createMainRuntimeRegistry();
  const overlayManager = params.createOverlayManager();
  const overlayModalInputState = params.createOverlayModalInputState({
    getModalWindow: () => overlayManager.getModalWindow(),
    syncOverlayShortcutsForModal: (isActive: boolean) => {
      params.getSyncOverlayShortcutsForModal()(isActive);
    },
    syncOverlayVisibilityForModal: () => {
      params.getSyncOverlayVisibilityForModal()();
    },
    restoreMainWindowFocus: () => {
      const mainWindow = overlayManager.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) return;
      try {
        const electron = require('electron') as {
          app?: { focus?: (options?: { steal?: boolean }) => void };
        };
        electron.app?.focus?.({ steal: true });
      } catch {
        // Ignore in non-Electron environments.
      }
      const maybeFocusable = mainWindow as typeof mainWindow & {
        setFocusable?: (focusable: boolean) => void;
      };
      maybeFocusable.setFocusable?.(true);
      mainWindow.focus();
      if (!mainWindow.webContents.isFocused()) {
        mainWindow.webContents.focus();
      }
    },
  });
  const overlayContentMeasurementStore = params.createOverlayContentMeasurementStore({
    logger,
  });
  const overlayModalRuntime = params.createOverlayModalRuntime({
    overlayManager,
    overlayModalInputState,
    onModalStateChange: (isActive: boolean) =>
      overlayModalInputState.handleModalInputStateChange(isActive),
  });
  const appState = params.createAppState({
    mpvSocketPath: params.getDefaultSocketPath(),
    texthookerPort: params.defaultTexthookerPort,
  });

  if (!params.existsSync(userDataPath)) {
    params.mkdirSync(userDataPath, { recursive: true });
  }
  params.app.setPath('userData', userDataPath);

  const appLifecycleApp = {
    requestSingleInstanceLock: () =>
      params.shouldBypassSingleInstanceLock() ? true : params.requestSingleInstanceLockEarly(),
    quit: () => params.app.quit(),
    exit: (code?: number) => params.app.exit(code),
    on: (event: string, listener: (...args: unknown[]) => void) => {
      if (event === 'second-instance') {
        params.registerSecondInstanceHandlerEarly(
          listener as (_event: unknown, argv: string[]) => void,
        );
        return appLifecycleApp;
      }
      params.app.on(event, listener);
      return appLifecycleApp;
    },
    whenReady: () => params.app.whenReady(),
  } satisfies AppLifecycleShape as TAppLifecycleApp;

  return {
    configDir,
    userDataPath,
    defaultMpvLogPath,
    defaultImmersionDbPath,
    configService,
    anilistTokenStore,
    jellyfinTokenStore,
    anilistUpdateQueue,
    subtitleWsService,
    annotationSubtitleWsService,
    logger,
    runtimeRegistry,
    overlayManager,
    overlayModalInputState,
    overlayContentMeasurementStore,
    overlayModalRuntime,
    appState,
    appLifecycleApp,
  };
}
