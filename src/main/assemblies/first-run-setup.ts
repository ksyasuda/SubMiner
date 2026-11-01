import os from 'node:os';
import type { BrowserWindow, BrowserWindowConstructorOptions } from 'electron';
import type { ConfigService } from '../../config/service';
import { getYomitanDictionaryInfo } from '../../core/services';
import { requestHachidoriSharing } from '../../core/services/tokenizer/yomitan-parser-runtime';
import type { Logger } from '../../logger';
import { resolveDefaultMpvInstallPaths } from '../../shared/setup-state';
import type { DictionaryBackend } from '../../types/config';
import {
  detectCommandLineLauncher,
  installBun as installCommandLineBun,
  installLauncher as installCommandLineLauncher,
} from '../runtime/command-line-launcher';
import {
  detectInstalledFirstRunPlugin,
  detectInstalledFirstRunPluginCandidates,
  removeLegacyMpvPluginCandidates,
} from '../runtime/first-run-setup-plugin';
import {
  createFirstRunSetupService,
  getFirstRunSetupCompletionMessage,
  isStandaloneFirstRunSetupCommand,
} from '../runtime/first-run-setup-service';
import {
  buildFirstRunSetupHtml,
  createMaybeFocusExistingFirstRunSetupWindowHandler,
  createOpenFirstRunSetupWindowHandler,
  parseFirstRunSetupSubmissionUrl,
  type FirstRunSetupSubmission,
} from '../runtime/first-run-setup-window';
import { createCreateFirstRunSetupWindowHandler } from '../runtime/setup-window-factory';
import { getConfiguredWindowsMpvPathStatus } from '../runtime/windows-mpv-launch';
import {
  applyWindowsMpvShortcuts,
  detectWindowsMpvShortcuts,
  resolveWindowsMpvShortcutPaths,
} from '../runtime/windows-mpv-shortcuts';
import { observeYomitanDictionaryCount } from '../runtime/yomitan-dictionary-integrity';
import type { AppState } from '../state';

type YomitanParserRuntimeDeps = Parameters<typeof requestHachidoriSharing>[1];
type CommandLineLauncherRuntimeOptions = NonNullable<
  Parameters<typeof detectCommandLineLauncher>[0]
>;

/**
 * What the first-run setup flow needs from the rest of the app. Siblings built after this
 * assembly in `main.ts` are reached through callbacks.
 */
export interface FirstRunSetupAssemblyDeps {
  appState: Pick<
    AppState,
    'firstRunSetupCompleted' | 'firstRunSetupWindow' | 'backgroundMode' | 'initialArgs'
  >;
  configService: Pick<ConfigService, 'getConfig' | 'patchRawConfig'>;
  logger: Pick<Logger, 'error' | 'info'>;
  configDir: string;
  userDataPath: string;
  /** Fixed for the process lifetime; backend changes apply on restart. */
  dictionaryBackend: DictionaryBackend;
  ensureYomitanExtensionLoaded: () => Promise<unknown>;
  getYomitanParserRuntimeDeps: () => YomitanParserRuntimeDeps;
  /** Resolves the launcher options after the login-shell PATH is applied. */
  createCommandLineLauncherRuntimeOptions: () => Promise<CommandLineLauncherRuntimeOptions>;
  getAppPath: (name: 'appData' | 'desktop') => string;
  trashItem: (targetPath: string) => Promise<void>;
  writeShortcutLink: Parameters<typeof applyWindowsMpvShortcuts>[0]['writeShortcutLink'];
  createBrowserWindow: (options: BrowserWindowConstructorOptions) => BrowserWindow;
  openDictionarySettings: () => boolean;
  openConfigSettingsWindow: () => boolean;
  /** Called when setup state changes, so the tray menu can refresh its setup entry. */
  onSetupStateChanged: () => void;
  requestAppQuit: () => void;
}

export type FirstRunSetupAssembly = ReturnType<typeof createFirstRunSetupAssembly>;

/**
 * Owns the first-run setup service and its window, plus the status message shown in that
 * window. `main.ts` initializes setup state during app-ready and opens the window from there,
 * the tray and the CLI.
 */
export function createFirstRunSetupAssembly(deps: FirstRunSetupAssemblyDeps) {
  const { appState, configService, logger } = deps;
  let firstRunSetupMessage: string | null = null;
  const resolveWindowsMpvShortcutRuntimePaths = () =>
    resolveWindowsMpvShortcutPaths({
      appDataDir: deps.getAppPath('appData'),
      desktopDir: deps.getAppPath('desktop'),
    });

  const firstRunSetupService = createFirstRunSetupService({
    getDictionaryBackend: () => deps.dictionaryBackend,
    getHachidoriHostStatus: async () => {
      await deps.ensureYomitanExtensionLoaded();
      return requestHachidoriSharing(
        { type: 'hd_sharing_status' },
        deps.getYomitanParserRuntimeDeps(),
        logger,
      );
    },
    platform: process.platform,
    configDir: deps.configDir,
    getYomitanDictionaryCount: async () => {
      await deps.ensureYomitanExtensionLoaded();
      const dictionaries = await getYomitanDictionaryInfo(deps.getYomitanParserRuntimeDeps(), {
        error: (message, ...args) => logger.error(message, ...args),
        info: (message, ...args) => logger.info(message, ...args),
      });
      const integrity = observeYomitanDictionaryCount(deps.userDataPath, dictionaries.length);
      if (!integrity.safe) {
        logger.error(`[dictionary:integrity] ${integrity.message}`);
      }
      return dictionaries.length;
    },
    isExternalYomitanConfigured: () =>
      configService.getConfig().yomitan.externalProfilePath.trim().length > 0,
    detectPluginInstalled: () => {
      const candidates = detectInstalledFirstRunPluginCandidates({
        platform: process.platform,
        homeDir: os.homedir(),
        xdgConfigHome: process.env.XDG_CONFIG_HOME,
        appDataDir: deps.getAppPath('appData'),
        mpvExecutablePath: configService.getConfig().mpv.executablePath,
      });
      if (candidates.length > 0) {
        return true;
      }
      const installPaths = resolveDefaultMpvInstallPaths(
        process.platform,
        os.homedir(),
        process.env.XDG_CONFIG_HOME,
      );
      return detectInstalledFirstRunPlugin(installPaths);
    },
    detectLegacyMpvPluginCandidates: () =>
      detectInstalledFirstRunPluginCandidates({
        platform: process.platform,
        homeDir: os.homedir(),
        xdgConfigHome: process.env.XDG_CONFIG_HOME,
        appDataDir: deps.getAppPath('appData'),
        mpvExecutablePath: configService.getConfig().mpv.executablePath,
      }),
    removeLegacyMpvPlugins: (candidates) =>
      removeLegacyMpvPluginCandidates({
        candidates,
        trashItem: (candidatePath) => deps.trashItem(candidatePath),
      }),
    detectWindowsMpvShortcuts: () => {
      if (process.platform !== 'win32') {
        return {
          startMenuInstalled: false,
          desktopInstalled: false,
        };
      }
      return detectWindowsMpvShortcuts(resolveWindowsMpvShortcutRuntimePaths());
    },
    applyWindowsMpvShortcuts: async (preferences) => {
      if (process.platform !== 'win32') {
        return {
          ok: true,
          status: 'unknown' as const,
          message: '',
        };
      }
      return applyWindowsMpvShortcuts({
        preferences,
        paths: resolveWindowsMpvShortcutRuntimePaths(),
        exePath: process.execPath,
        writeShortcutLink: (shortcutPath, operation, details) =>
          deps.writeShortcutLink(shortcutPath, operation, details),
      });
    },
    detectCommandLineLauncher: async () =>
      detectCommandLineLauncher(await deps.createCommandLineLauncherRuntimeOptions()),
    installBun: async () => {
      const snapshot = await installCommandLineBun(
        await deps.createCommandLineLauncherRuntimeOptions(),
      );
      return {
        ok: snapshot.status === 'ready',
        message:
          snapshot.message ??
          (snapshot.status === 'ready'
            ? 'Bun is ready. Open a new terminal.'
            : 'Bun installation failed.'),
      };
    },
    installCommandLineLauncher: async () => {
      const snapshot = await installCommandLineLauncher(
        await deps.createCommandLineLauncherRuntimeOptions(),
      );
      const ok = snapshot.status === 'ready' || snapshot.status === 'not_on_path';
      return {
        ok,
        installPath: snapshot.installPath,
        message:
          snapshot.message ??
          (ok
            ? 'Command-line launcher installed. Open a new terminal.'
            : 'Command-line launcher installation failed.'),
      };
    },
    onStateChanged: (state) => {
      appState.firstRunSetupCompleted = state.status === 'completed';
      deps.onSetupStateChanged();
    },
  });

  const maybeFocusExistingFirstRunSetupWindow = createMaybeFocusExistingFirstRunSetupWindowHandler({
    getSetupWindow: () => appState.firstRunSetupWindow,
  });
  const openFirstRunSetupWindowHandler = createOpenFirstRunSetupWindowHandler({
    maybeFocusExistingSetupWindow: maybeFocusExistingFirstRunSetupWindow,
    createSetupWindow: createCreateFirstRunSetupWindowHandler({
      createBrowserWindow: (options) => deps.createBrowserWindow(options),
    }),
    getSetupSnapshot: async () => {
      const snapshot = await firstRunSetupService.getSetupStatus();
      const mpvExecutablePath = configService.getConfig().mpv.executablePath;
      return {
        configReady: snapshot.configReady,
        dictionaryCount: snapshot.dictionaryCount,
        dictionaryBackend: snapshot.dictionaryBackend,
        hachidoriHost: snapshot.hachidoriHost,
        canFinish: snapshot.canFinish,
        externalYomitanConfigured: snapshot.externalYomitanConfigured,
        pluginStatus: snapshot.pluginStatus,
        pluginInstallPathSummary: snapshot.pluginInstallPathSummary,
        legacyMpvPluginPaths: snapshot.legacyMpvPluginPaths,
        mpvExecutablePath,
        mpvExecutablePathStatus: getConfiguredWindowsMpvPathStatus(mpvExecutablePath),
        windowsMpvShortcuts: snapshot.windowsMpvShortcuts,
        commandLineLauncher: snapshot.commandLineLauncher,
        message: firstRunSetupMessage,
      };
    },
    buildSetupHtml: (model) => buildFirstRunSetupHtml(model),
    parseSubmissionUrl: (rawUrl) => parseFirstRunSetupSubmissionUrl(rawUrl),
    handleAction: async (submission: FirstRunSetupSubmission) => {
      if (
        submission.action === 'link-hachidori-host' ||
        submission.action === 'unlink-hachidori-host'
      ) {
        try {
          if (deps.dictionaryBackend !== 'hachidori')
            throw new Error('Select Hachidori and restart SubMiner before linking a host.');
          await deps.ensureYomitanExtensionLoaded();
          const status = await requestHachidoriSharing(
            submission.action === 'link-hachidori-host'
              ? { type: 'hd_sharing_client_link', address: submission.address }
              : { type: 'hd_sharing_client_unlink' },
            deps.getYomitanParserRuntimeDeps(),
            logger,
          );
          firstRunSetupMessage =
            status.kind === 'local'
              ? 'Using dictionaries installed in SubMiner.'
              : status.kind === 'connected'
                ? `Linked to ${status.name}. Anki mining stays in SubMiner.`
                : status.message;
        } catch (error) {
          firstRunSetupMessage =
            error instanceof Error ? error.message : 'Could not update the dictionary host.';
        }
        return;
      }
      if (submission.action === 'remove-legacy-plugin') {
        const snapshot = await firstRunSetupService.removeLegacyMpvPlugin();
        firstRunSetupMessage = snapshot.message;
        return;
      }
      if (submission.action === 'configure-mpv-executable-path') {
        const mpvExecutablePath = submission.mpvExecutablePath?.trim() ?? '';
        const pathStatus = getConfiguredWindowsMpvPathStatus(mpvExecutablePath);
        configService.patchRawConfig({
          mpv: {
            executablePath: mpvExecutablePath,
          },
        });
        firstRunSetupMessage =
          pathStatus === 'invalid'
            ? `Saved mpv executable path, but the file was not found: ${mpvExecutablePath}`
            : mpvExecutablePath
              ? `Saved mpv executable path: ${mpvExecutablePath}`
              : 'Cleared mpv executable path. SubMiner will auto-discover mpv.exe from SUBMINER_MPV_PATH or PATH.';
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
      if (submission.action === 'install-bun') {
        const snapshot = await firstRunSetupService.installBun();
        firstRunSetupMessage = snapshot.message;
        return;
      }
      if (submission.action === 'install-command-line-launcher') {
        const snapshot = await firstRunSetupService.installCommandLineLauncher();
        firstRunSetupMessage = snapshot.message;
        return;
      }
      if (submission.action === 'open-yomitan-settings') {
        firstRunSetupMessage = deps.openDictionarySettings()
          ? `Opened ${deps.dictionaryBackend === 'hachidori' ? 'Hachidori' : 'Yomitan'} settings. Install dictionaries, then refresh status.`
          : 'Yomitan settings are unavailable while external read-only profile mode is enabled.';
        return;
      }
      if (submission.action === 'open-config-settings') {
        const opened = deps.openConfigSettingsWindow();
        firstRunSetupMessage = opened
          ? 'Opened SubMiner settings.'
          : 'SubMiner settings are unavailable.';
        if (opened) {
          return { skipRender: true };
        }
        return;
      }
      if (submission.action === 'refresh') {
        const snapshot = await firstRunSetupService.refreshStatus('Status refreshed.');
        firstRunSetupMessage = snapshot.message;
        return;
      }

      const snapshot = await firstRunSetupService.markSetupCompleted();
      if (snapshot.state.status === 'completed') {
        firstRunSetupMessage = null;
        return { closeWindow: true };
      }
      firstRunSetupMessage =
        getFirstRunSetupCompletionMessage(snapshot) ??
        'Finish setup requires the mpv plugin and Yomitan dictionaries.';
      return;
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
    shouldQuitWhenClosedIncomplete: () => !appState.backgroundMode,
    shouldQuitWhenClosedCompleted: () =>
      Boolean(appState.initialArgs && isStandaloneFirstRunSetupCommand(appState.initialArgs)),
    quitApp: () => deps.requestAppQuit(),
    clearSetupWindow: () => {
      appState.firstRunSetupWindow = null;
    },
    setSetupWindow: (window) => {
      appState.firstRunSetupWindow = window as BrowserWindow;
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
    logError: (message, error) => logger.error(message, error),
  });

  /** Opens the setup window; without `force` it is a no-op once setup has completed. */
  function openWindow(force = false): void {
    if (!force && firstRunSetupService.isSetupCompleted()) {
      return;
    }
    openFirstRunSetupWindowHandler();
  }

  return { service: firstRunSetupService, openWindow };
}
