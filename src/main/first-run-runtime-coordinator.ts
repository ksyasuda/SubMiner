import { BrowserWindow } from 'electron';

import { getYomitanDictionaryInfo } from '../core/services';
import type { ResolvedConfig } from '../types';
import { createFirstRunRuntime } from './first-run-runtime';
import type { AppState } from './state';

export interface FirstRunRuntimeCoordinatorInput {
  platform: NodeJS.Platform;
  configDir: string;
  homeDir: string;
  xdgConfigHome?: string;
  binaryPath: string;
  appPath: string;
  resourcesPath: string;
  appDataDir: string;
  desktopDir: string;
  appState: Pick<AppState, 'firstRunSetupWindow' | 'firstRunSetupCompleted' | 'backgroundMode'>;
  getResolvedConfig: () => ResolvedConfig;
  yomitan: {
    ensureYomitanExtensionLoaded: () => Promise<unknown>;
    getParserRuntimeDeps: () => Parameters<typeof getYomitanDictionaryInfo>[0];
    openYomitanSettings: () => boolean;
  };
  overlay: {
    ensureTray: () => void;
    hasTray: () => boolean;
  };
  actions: {
    writeShortcutLink: (
      shortcutPath: string,
      operation: 'create' | 'update' | 'replace',
      details: {
        target: string;
        args?: string;
        cwd?: string;
        description?: string;
        icon?: string;
        iconIndex?: number;
      },
    ) => boolean;
    requestAppQuit: () => void;
  };
  logger: {
    error: (message: string, error: unknown) => void;
    info: (message: string, ...args: unknown[]) => void;
  };
}

export function createFirstRunRuntimeCoordinator(input: FirstRunRuntimeCoordinatorInput) {
  return createFirstRunRuntime<BrowserWindow>({
    platform: input.platform,
    configDir: input.configDir,
    homeDir: input.homeDir,
    xdgConfigHome: input.xdgConfigHome,
    binaryPath: input.binaryPath,
    appPath: input.appPath,
    resourcesPath: input.resourcesPath,
    appDataDir: input.appDataDir,
    desktopDir: input.desktopDir,
    getYomitanDictionaryCount: async () => {
      await input.yomitan.ensureYomitanExtensionLoaded();
      const dictionaries = await getYomitanDictionaryInfo(input.yomitan.getParserRuntimeDeps(), {
        error: (message, ...args) => input.logger.error(message, args[0]),
        info: (message, ...args) => input.logger.info(message, ...args),
      });
      return dictionaries.length;
    },
    isExternalYomitanConfigured: () =>
      input.getResolvedConfig().yomitan.externalProfilePath.trim().length > 0,
    createBrowserWindow: (options) => {
      const window = new BrowserWindow(options);
      input.appState.firstRunSetupWindow = window;
      window.on('closed', () => {
        input.appState.firstRunSetupWindow = null;
      });
      return window;
    },
    writeShortcutLink: (shortcutPath, operation, details) =>
      input.actions.writeShortcutLink(shortcutPath, operation, details),
    openYomitanSettings: () => input.yomitan.openYomitanSettings(),
    shouldQuitWhenClosedIncomplete: () => !input.appState.backgroundMode,
    quitApp: () => input.actions.requestAppQuit(),
    logError: (message, error) => input.logger.error(message, error),
    onStateChanged: (state) => {
      input.appState.firstRunSetupCompleted = state.status === 'completed';
      if (input.overlay.hasTray()) {
        input.overlay.ensureTray();
      }
    },
  });
}
