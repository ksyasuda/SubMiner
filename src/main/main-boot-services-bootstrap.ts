import type { BrowserWindow } from 'electron';

import type { ConfigStartupParseError } from '../config';
import {
  createMainBootServices,
  type MainBootServicesResult,
  type OverlayModalInputStateShape,
  type AppLifecycleShape,
} from './boot/services';

export interface MainBootServicesBootstrapInput<
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
  system: {
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
      // eslint-disable-next-line @typescript-eslint/no-unsafe-function-type -- Electron App.on has 50+ overloaded signatures
      on: Function;
      whenReady: () => Promise<void>;
    };
  };
  singleInstance: {
    shouldBypassSingleInstanceLock: () => boolean;
    requestSingleInstanceLockEarly: () => boolean;
    registerSecondInstanceHandlerEarly: (
      listener: (_event: unknown, argv: string[]) => void,
    ) => void;
    onConfigStartupParseError: (error: ConfigStartupParseError) => void;
  };
  factories: {
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
  };
}

export function createMainBootServicesBootstrap<
  TConfigService,
  TAnilistTokenStore,
  TJellyfinTokenStore,
  TAnilistUpdateQueue,
  TSubtitleWebSocket,
  TLogger,
  TRuntimeRegistry,
  TOverlayManager extends { getModalWindow: () => BrowserWindow | null },
  TOverlayModalInputState extends OverlayModalInputStateShape,
  TOverlayContentMeasurementStore,
  TOverlayModalRuntime,
  TAppState,
  TAppLifecycleApp extends AppLifecycleShape,
>(
  input: MainBootServicesBootstrapInput<
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
  return createMainBootServices({
    platform: input.system.platform,
    argv: input.system.argv,
    appDataDir: input.system.appDataDir,
    xdgConfigHome: input.system.xdgConfigHome,
    homeDir: input.system.homeDir,
    defaultMpvLogFile: input.system.defaultMpvLogFile,
    envMpvLog: input.system.envMpvLog,
    defaultTexthookerPort: input.system.defaultTexthookerPort,
    getDefaultSocketPath: input.system.getDefaultSocketPath,
    resolveConfigDir: input.system.resolveConfigDir,
    existsSync: input.system.existsSync,
    mkdirSync: input.system.mkdirSync,
    joinPath: input.system.joinPath,
    app: input.system.app,
    shouldBypassSingleInstanceLock: input.singleInstance.shouldBypassSingleInstanceLock,
    requestSingleInstanceLockEarly: input.singleInstance.requestSingleInstanceLockEarly,
    registerSecondInstanceHandlerEarly: input.singleInstance.registerSecondInstanceHandlerEarly,
    onConfigStartupParseError: input.singleInstance.onConfigStartupParseError,
    createConfigService: input.factories.createConfigService,
    createAnilistTokenStore: input.factories.createAnilistTokenStore,
    createJellyfinTokenStore: input.factories.createJellyfinTokenStore,
    createAnilistUpdateQueue: input.factories.createAnilistUpdateQueue,
    createSubtitleWebSocket: input.factories.createSubtitleWebSocket,
    createLogger: input.factories.createLogger,
    createMainRuntimeRegistry: input.factories.createMainRuntimeRegistry,
    createOverlayManager: input.factories.createOverlayManager,
    createOverlayModalInputState: input.factories.createOverlayModalInputState,
    createOverlayContentMeasurementStore: input.factories.createOverlayContentMeasurementStore,
    getSyncOverlayShortcutsForModal: input.factories.getSyncOverlayShortcutsForModal,
    getSyncOverlayVisibilityForModal: input.factories.getSyncOverlayVisibilityForModal,
    createOverlayModalRuntime: input.factories.createOverlayModalRuntime,
    createAppState: input.factories.createAppState,
  });
}
