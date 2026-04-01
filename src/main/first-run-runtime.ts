import {
  createFirstRunSetupService,
  shouldAutoOpenFirstRunSetup,
  type FirstRunSetupService,
  type PluginInstallResult,
  type SetupStatusSnapshot,
} from './runtime/first-run-setup-service';
import type { SetupState } from '../shared/setup-state';
import {
  buildFirstRunSetupHtml,
  createMaybeFocusExistingFirstRunSetupWindowHandler,
  createOpenFirstRunSetupWindowHandler,
  parseFirstRunSetupSubmissionUrl,
} from './runtime/first-run-setup-window';
import { createCreateFirstRunSetupWindowHandler } from './runtime/setup-window-factory';
import {
  detectInstalledFirstRunPlugin,
  installFirstRunPluginToDefaultLocation,
  syncInstalledFirstRunPluginBinaryPath,
} from './runtime/first-run-setup-plugin';
import {
  applyWindowsMpvShortcuts,
  detectWindowsMpvShortcuts,
  resolveWindowsMpvShortcutPaths,
} from './runtime/windows-mpv-shortcuts';
import { resolveDefaultMpvInstallPaths } from '../shared/setup-state';

export interface FirstRunSetupWindowLike {
  webContents: {
    on: (event: 'will-navigate', handler: (event: unknown, url: string) => void) => void;
  };
  loadURL: (url: string) => Promise<void> | void;
  on: (event: 'closed', handler: () => void) => void;
  isDestroyed: () => boolean;
  close: () => void;
  focus: () => void;
}

export interface FirstRunRuntimeInput<
  TWindow extends FirstRunSetupWindowLike = FirstRunSetupWindowLike,
> {
  platform: NodeJS.Platform;
  configDir: string;
  homeDir: string;
  xdgConfigHome?: string;
  binaryPath: string;
  appPath: string;
  resourcesPath: string;
  appDataDir: string;
  desktopDir: string;
  getYomitanDictionaryCount: () => Promise<number>;
  isExternalYomitanConfigured: () => boolean;
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
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
  openYomitanSettings: () => boolean;
  shouldQuitWhenClosedIncomplete: () => boolean;
  quitApp: () => void;
  logError: (message: string, error: unknown) => void;
  onStateChanged?: (state: SetupState) => void;
}

export interface FirstRunRuntime {
  ensureSetupStateInitialized: () => Promise<SetupStatusSnapshot>;
  isSetupCompleted: () => boolean;
  openFirstRunSetupWindow: () => void;
}

export function createFirstRunRuntime<TWindow extends FirstRunSetupWindowLike>(
  input: FirstRunRuntimeInput<TWindow>,
): FirstRunRuntime {
  syncInstalledFirstRunPluginBinaryPath({
    platform: input.platform,
    homeDir: input.homeDir,
    xdgConfigHome: input.xdgConfigHome,
    binaryPath: input.binaryPath,
  });

  const firstRunSetupService = createFirstRunSetupService({
    platform: input.platform,
    configDir: input.configDir,
    getYomitanDictionaryCount: input.getYomitanDictionaryCount,
    isExternalYomitanConfigured: input.isExternalYomitanConfigured,
    detectPluginInstalled: () =>
      detectInstalledFirstRunPlugin(
        resolveDefaultMpvInstallPaths(input.platform, input.homeDir, input.xdgConfigHome),
      ),
    installPlugin: async (): Promise<PluginInstallResult> =>
      installFirstRunPluginToDefaultLocation({
        platform: input.platform,
        homeDir: input.homeDir,
        xdgConfigHome: input.xdgConfigHome,
        dirname: __dirname,
        appPath: input.appPath,
        resourcesPath: input.resourcesPath,
        binaryPath: input.binaryPath,
      }),
    detectWindowsMpvShortcuts: async () =>
      detectWindowsMpvShortcuts(
        resolveWindowsMpvShortcutPaths({
          appDataDir: input.appDataDir,
          desktopDir: input.desktopDir,
        }),
      ),
    applyWindowsMpvShortcuts: async (preferences) =>
      applyWindowsMpvShortcuts({
        preferences,
        paths: resolveWindowsMpvShortcutPaths({
          appDataDir: input.appDataDir,
          desktopDir: input.desktopDir,
        }),
        exePath: input.binaryPath,
        writeShortcutLink: (shortcutPath, operation, details) =>
          input.writeShortcutLink(shortcutPath, operation, details),
      }),
    onStateChanged: (state) => {
      input.onStateChanged?.(state);
    },
  });

  let firstRunSetupWindow: TWindow | null = null;
  let firstRunSetupMessage: string | null = null;

  const maybeFocusExistingFirstRunSetupWindow = createMaybeFocusExistingFirstRunSetupWindowHandler({
    getSetupWindow: () => firstRunSetupWindow,
  });

  const createSetupWindow = createCreateFirstRunSetupWindowHandler({
    createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) =>
      input.createBrowserWindow(options),
  });

  const openFirstRunSetupWindowHandler = createOpenFirstRunSetupWindowHandler({
    maybeFocusExistingSetupWindow: () => maybeFocusExistingFirstRunSetupWindow(),
    createSetupWindow: () => {
      const window = createSetupWindow();
      firstRunSetupWindow = window;
      return window;
    },
    getSetupSnapshot: async () => {
      const snapshot = await firstRunSetupService.getSetupStatus();
      return {
        ...snapshot,
        message: firstRunSetupMessage,
      };
    },
    buildSetupHtml: (model) => buildFirstRunSetupHtml(model),
    parseSubmissionUrl: (rawUrl) => parseFirstRunSetupSubmissionUrl(rawUrl),
    handleAction: async (submission) => {
      if (submission.action === 'install-plugin') {
        const snapshot = await firstRunSetupService.installMpvPlugin();
        firstRunSetupMessage = snapshot.message;
        return;
      }

      if (submission.action === 'configure-windows-mpv-shortcuts') {
        const snapshot = await firstRunSetupService.configureWindowsMpvShortcuts({
          startMenuEnabled: submission.startMenuEnabled === true,
          desktopEnabled: submission.desktopEnabled === true,
        });
        firstRunSetupMessage = snapshot.message;
        return;
      }

      if (submission.action === 'open-yomitan-settings') {
        firstRunSetupMessage = input.openYomitanSettings()
          ? 'Opened Yomitan settings. Install dictionaries, then refresh status.'
          : 'Yomitan settings are unavailable while external read-only profile mode is enabled.';
        return;
      }

      if (submission.action === 'refresh') {
        const snapshot = await firstRunSetupService.refreshStatus('Status refreshed.');
        firstRunSetupMessage = snapshot.message;
        return;
      }

      if (submission.action === 'skip-plugin') {
        await firstRunSetupService.skipPluginInstall();
        firstRunSetupMessage = 'mpv plugin installation skipped.';
        return;
      }

      const snapshot = await firstRunSetupService.markSetupCompleted();
      if (snapshot.state.status === 'completed') {
        firstRunSetupMessage = null;
        return { closeWindow: true };
      }
      firstRunSetupMessage = 'Install at least one Yomitan dictionary before finishing setup.';
      return undefined;
    },
    markSetupInProgress: async () => {
      firstRunSetupMessage = null;
      await firstRunSetupService.markSetupInProgress();
    },
    markSetupCancelled: async () => {
      firstRunSetupMessage = null;
      await firstRunSetupService.markSetupCancelled();
    },
    isSetupCompleted: () => firstRunSetupService.isSetupCompleted(),
    shouldQuitWhenClosedIncomplete: () => input.shouldQuitWhenClosedIncomplete(),
    quitApp: () => input.quitApp(),
    clearSetupWindow: () => {
      firstRunSetupWindow = null;
    },
    setSetupWindow: (window) => {
      firstRunSetupWindow = window;
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
    logError: (message, error) => input.logError(message, error),
  });

  return {
    ensureSetupStateInitialized: () => firstRunSetupService.ensureSetupStateInitialized(),
    isSetupCompleted: () => firstRunSetupService.isSetupCompleted(),
    openFirstRunSetupWindow: () => {
      if (firstRunSetupService.isSetupCompleted()) {
        return;
      }
      openFirstRunSetupWindowHandler();
    },
  };
}

export { shouldAutoOpenFirstRunSetup };
