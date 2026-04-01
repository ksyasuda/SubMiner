import * as fs from 'node:fs';
import * as path from 'node:path';
import type { BrowserWindow } from 'electron';

import { createAnilistTokenStore } from '../core/services/anilist/anilist-token-store';
import { createJellyfinTokenStore } from '../core/services/jellyfin-token-store';
import { createAnilistUpdateQueue } from '../core/services/anilist/anilist-update-queue';
import {
  SubtitleWebSocket,
  createOverlayContentMeasurementStore,
  createOverlayManager,
} from '../core/services';
import { ConfigService } from '../config';
import { resolveConfigDir } from '../config/path-resolution';
import { createAppState } from './state';
import {
  createMainBootServices,
  type AppLifecycleShape,
  type MainBootServicesResult,
} from './boot/services';
import { createLogger } from '../logger';
import { createMainRuntimeRegistry } from './runtime/registry';
import { createOverlayModalInputState } from './runtime/overlay-modal-input-state';
import { createOverlayModalRuntimeService } from './overlay-runtime';
import { buildConfigParseErrorDetails, failStartupFromConfig } from './config-validation';
import {
  registerSecondInstanceHandlerEarly,
  requestSingleInstanceLockEarly,
  shouldBypassSingleInstanceLockForArgv,
} from './early-single-instance';
import {
  createBuildOverlayContentMeasurementStoreMainDepsHandler,
  createBuildOverlayModalRuntimeMainDepsHandler,
} from './runtime/domains/overlay';
import type { WindowGeometry } from '../types';

export type MainBootRuntime = MainBootServicesResult<
  ConfigService,
  ReturnType<typeof createAnilistTokenStore>,
  ReturnType<typeof createJellyfinTokenStore>,
  ReturnType<typeof createAnilistUpdateQueue>,
  SubtitleWebSocket,
  ReturnType<typeof createLogger>,
  ReturnType<typeof createMainRuntimeRegistry>,
  ReturnType<typeof createOverlayManager>,
  ReturnType<typeof createOverlayModalInputState>,
  ReturnType<typeof createOverlayContentMeasurementStore>,
  ReturnType<typeof createOverlayModalRuntimeService>,
  ReturnType<typeof createAppState>,
  AppLifecycleShape
>;

export interface MainBootRuntimeInput {
  platform: NodeJS.Platform;
  argv: string[];
  appDataDir: string | undefined;
  xdgConfigHome: string | undefined;
  homeDir: string;
  defaultMpvLogFile: string;
  envMpvLog: string | undefined;
  defaultTexthookerPort: number;
  getDefaultSocketPath: () => string;
  app: {
    setPath: (name: string, value: string) => void;
    quit: () => void;
    // eslint-disable-next-line @typescript-eslint/no-unsafe-function-type -- Electron App.on has 50+ overloaded signatures
    on: Function;
    whenReady: () => Promise<void>;
  };
  dialog: {
    showErrorBox: (title: string, details: string) => void;
  };
  overlay: {
    getSyncOverlayShortcutsForModal: () => (isActive: boolean) => void;
    getSyncOverlayVisibilityForModal: () => () => void;
    createModalWindow: () => BrowserWindow;
    getOverlayGeometry: () => WindowGeometry;
  };
  notifications: {
    notifyAnilistTokenStoreWarning: (message: string) => void;
    requestAppQuit: () => void;
  };
}

export function createMainBootRuntime(input: MainBootRuntimeInput): MainBootRuntime {
  return createMainBootServices({
    platform: input.platform,
    argv: input.argv,
    appDataDir: input.appDataDir,
    xdgConfigHome: input.xdgConfigHome,
    homeDir: input.homeDir,
    defaultMpvLogFile: input.defaultMpvLogFile,
    envMpvLog: input.envMpvLog,
    defaultTexthookerPort: input.defaultTexthookerPort,
    getDefaultSocketPath: () => input.getDefaultSocketPath(),
    resolveConfigDir,
    existsSync: (targetPath) => fs.existsSync(targetPath),
    mkdirSync: (targetPath, options) => {
      fs.mkdirSync(targetPath, options);
    },
    joinPath: (...parts) => path.join(...parts),
    app: input.app,
    shouldBypassSingleInstanceLock: () => shouldBypassSingleInstanceLockForArgv(input.argv),
    requestSingleInstanceLockEarly: () => requestSingleInstanceLockEarly(input.app as never),
    registerSecondInstanceHandlerEarly: (listener) => {
      registerSecondInstanceHandlerEarly(input.app as never, listener);
    },
    onConfigStartupParseError: (error) => {
      failStartupFromConfig(
        'SubMiner config parse error',
        buildConfigParseErrorDetails(error.path, error.parseError),
        {
          logError: (details) => console.error(details),
          showErrorBox: (title, details) => input.dialog.showErrorBox(title, details),
          quit: () => input.notifications.requestAppQuit(),
        },
      );
    },
    createConfigService: (configDir) => new ConfigService(configDir),
    createAnilistTokenStore: (targetPath) =>
      createAnilistTokenStore(targetPath, {
        info: (message: string) => console.info(message),
        warn: (message: string, details?: unknown) => console.warn(message, details),
        error: (message: string, details?: unknown) => console.error(message, details),
        warnUser: (message: string) => input.notifications.notifyAnilistTokenStoreWarning(message),
      }),
    createJellyfinTokenStore: (targetPath) =>
      createJellyfinTokenStore(targetPath, {
        info: (message: string) => console.info(message),
        warn: (message: string, details?: unknown) => console.warn(message, details),
        error: (message: string, details?: unknown) => console.error(message, details),
      }),
    createAnilistUpdateQueue: (targetPath) =>
      createAnilistUpdateQueue(targetPath, {
        info: (message: string) => console.info(message),
        warn: (message: string, details?: unknown) => console.warn(message, details),
        error: (message: string, details?: unknown) => console.error(message, details),
      }),
    createSubtitleWebSocket: () => new SubtitleWebSocket(),
    createLogger,
    createMainRuntimeRegistry,
    createOverlayManager,
    createOverlayModalInputState,
    createOverlayContentMeasurementStore: ({ logger }) =>
      createOverlayContentMeasurementStore(
        createBuildOverlayContentMeasurementStoreMainDepsHandler({
          now: () => Date.now(),
          warn: (message: string) => logger.warn(message),
        })(),
      ),
    getSyncOverlayShortcutsForModal: () => input.overlay.getSyncOverlayShortcutsForModal(),
    getSyncOverlayVisibilityForModal: () => input.overlay.getSyncOverlayVisibilityForModal(),
    createOverlayModalRuntime: ({ overlayManager, overlayModalInputState }) =>
      createOverlayModalRuntimeService(
        createBuildOverlayModalRuntimeMainDepsHandler({
          getMainWindow: () => overlayManager.getMainWindow(),
          getModalWindow: () => overlayManager.getModalWindow(),
          createModalWindow: () => input.overlay.createModalWindow(),
          getModalGeometry: () => input.overlay.getOverlayGeometry(),
          setModalWindowBounds: (geometry) => overlayManager.setModalWindowBounds(geometry),
        })(),
        {
          onModalStateChange: (isActive: boolean) =>
            overlayModalInputState.handleModalInputStateChange(isActive),
        },
      ),
    createAppState,
  }) as MainBootRuntime;
}
