import { createConfigHotReloadRuntime, type MpvIpcClient } from '../core/services';
import type { MpvRuntimeClientLike } from '../core/services/mpv';
import type {
  ConfigHotReloadPayload,
  ConfigValidationWarning,
  JimakuLanguagePreference,
  ResolvedConfig,
  SecondarySubMode,
  SubsyncManualPayload,
} from '../types';
import type { ReloadConfigStrictResult } from '../config';
import { RuntimeOptionsManager } from '../runtime-options';
import {
  createApplyJellyfinMpvDefaultsHandler,
  createBuildApplyJellyfinMpvDefaultsMainDepsHandler,
  createBuildGetDefaultSocketPathMainDepsHandler,
  createGetDefaultSocketPathHandler,
} from './runtime/domains/jellyfin';
import {
  createBuildConfigHotReloadAppliedMainDepsHandler,
  createBuildConfigHotReloadMessageMainDepsHandler,
  createBuildConfigHotReloadRuntimeMainDepsHandler,
  createBuildWatchConfigPathMainDepsHandler,
  createConfigHotReloadAppliedHandler,
  createConfigHotReloadMessageHandler,
  createWatchConfigPathHandler,
  buildRestartRequiredConfigMessage,
} from './runtime/domains/overlay';
import {
  createBuildConfigDerivedRuntimeMainDepsHandler,
  createBuildImmersionMediaRuntimeMainDepsHandler,
  createBuildMainSubsyncRuntimeMainDepsHandler,
  createConfigDerivedRuntime,
  createImmersionMediaRuntime,
  createMainSubsyncRuntime,
} from './runtime/domains/startup';
import {
  buildConfigWarningDialogDetails,
  buildConfigWarningNotificationBody,
} from './config-validation';

type ImmersionTrackerLike = {
  handleMediaChange: (path: string, title: string | null) => void;
};

type MpvClientLike = MpvIpcClient | null;
type JellyfinMpvClientLike = MpvRuntimeClientLike;

export interface StartupSupportRuntimeInput {
  platform: NodeJS.Platform;
  defaultImmersionDbPath: string;
  defaultJimakuLanguagePreference: JimakuLanguagePreference;
  defaultJimakuMaxEntryResults: number;
  defaultJimakuApiBaseUrl: string;
  jellyfinLangPref: string;
  getResolvedConfig: () => ResolvedConfig;
  appState: {
    immersionTracker: ImmersionTrackerLike | null;
    mpvClient: MpvClientLike;
    currentMediaPath: string | null;
    currentMediaTitle: string | null;
    runtimeOptionsManager: RuntimeOptionsManager | null;
    subsyncInProgress: boolean;
    keybindings: ConfigHotReloadPayload['keybindings'];
    ankiIntegration: {
      applyRuntimeConfigPatch: (patch: {
        ai: ResolvedConfig['ankiConnect']['ai']['enabled'];
      }) => void;
    } | null;
  };
  mpv: {
    sendMpvCommandRuntime: (client: JellyfinMpvClientLike, command: (string | number)[]) => void;
    showMpvOsd: (text: string) => void;
  };
  config: {
    reloadConfigStrict: () => ReloadConfigStrictResult;
  };
  subsync: {
    openManualPicker: (payload: SubsyncManualPayload) => void;
  };
  hotReload: {
    setSecondarySubMode: (mode: SecondarySubMode) => void;
    refreshGlobalAndOverlayShortcuts: () => void;
    broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  };
  notifications: {
    showDesktopNotification: (title: string, options: { body?: string }) => void;
    showErrorBox: (title: string, details: string) => void;
  };
  logger: {
    debug: (message: string) => void;
    info: (message: string) => void;
    warn: (message: string, error?: unknown) => void;
  };
  watch: {
    fileExists: (targetPath: string) => boolean;
    dirname: (targetPath: string) => string;
    watchPath: (
      targetPath: string,
      listener: (eventType: string, filename: string | null) => void,
    ) => { close: () => void };
  };
  timers: {
    setTimeout: (callback: () => void, delayMs: number) => NodeJS.Timeout;
    clearTimeout: (timeout: NodeJS.Timeout) => void;
  };
}

export interface StartupSupportRuntime {
  applyJellyfinMpvDefaults: (client: JellyfinMpvClientLike) => void;
  getDefaultSocketPath: () => string;
  immersionMediaRuntime: ReturnType<typeof createImmersionMediaRuntime>;
  configDerivedRuntime: ReturnType<typeof createConfigDerivedRuntime>;
  subsyncRuntime: ReturnType<typeof createMainSubsyncRuntime>;
  configHotReloadRuntime: ReturnType<typeof createConfigHotReloadRuntime>;
}

export function createStartupSupportRuntime(
  input: StartupSupportRuntimeInput,
): StartupSupportRuntime {
  const applyJellyfinMpvDefaultsHandler = createApplyJellyfinMpvDefaultsHandler(
    createBuildApplyJellyfinMpvDefaultsMainDepsHandler({
      sendMpvCommandRuntime: (client, command) => input.mpv.sendMpvCommandRuntime(client, command),
      jellyfinLangPref: input.jellyfinLangPref,
    })(),
  );

  const getDefaultSocketPathHandler = createGetDefaultSocketPathHandler(
    createBuildGetDefaultSocketPathMainDepsHandler({
      platform: input.platform,
    })(),
  );

  const immersionMediaRuntime = createImmersionMediaRuntime(
    createBuildImmersionMediaRuntimeMainDepsHandler({
      getResolvedConfig: () => input.getResolvedConfig(),
      defaultImmersionDbPath: input.defaultImmersionDbPath,
      getTracker: () => input.appState.immersionTracker,
      getMpvClient: () => input.appState.mpvClient,
      getCurrentMediaPath: () => input.appState.currentMediaPath,
      getCurrentMediaTitle: () => input.appState.currentMediaTitle,
      logDebug: (message) => input.logger.debug(message),
      logInfo: (message) => input.logger.info(message),
    })(),
  );

  const configDerivedRuntime = createConfigDerivedRuntime(
    createBuildConfigDerivedRuntimeMainDepsHandler({
      getResolvedConfig: () => input.getResolvedConfig(),
      getRuntimeOptionsManager: () => input.appState.runtimeOptionsManager,
      defaultJimakuLanguagePreference: input.defaultJimakuLanguagePreference,
      defaultJimakuMaxEntryResults: input.defaultJimakuMaxEntryResults,
      defaultJimakuApiBaseUrl: input.defaultJimakuApiBaseUrl,
    })(),
  );

  const subsyncRuntime = createMainSubsyncRuntime(
    createBuildMainSubsyncRuntimeMainDepsHandler({
      getMpvClient: () => input.appState.mpvClient,
      getResolvedConfig: () => input.getResolvedConfig(),
      getSubsyncInProgress: () => input.appState.subsyncInProgress,
      setSubsyncInProgress: (inProgress) => {
        input.appState.subsyncInProgress = inProgress;
      },
      showMpvOsd: (text) => input.mpv.showMpvOsd(text),
      openManualPicker: (payload) => input.subsync.openManualPicker(payload),
    })(),
  );

  const notifyConfigHotReloadMessage = createConfigHotReloadMessageHandler(
    createBuildConfigHotReloadMessageMainDepsHandler({
      showMpvOsd: (message) => input.mpv.showMpvOsd(message),
      showDesktopNotification: (title, options) =>
        input.notifications.showDesktopNotification(title, options),
    })(),
  );

  const watchConfigPathHandler = createWatchConfigPathHandler(
    createBuildWatchConfigPathMainDepsHandler({
      fileExists: (targetPath) => input.watch.fileExists(targetPath),
      dirname: (targetPath) => input.watch.dirname(targetPath),
      watchPath: (targetPath, listener) => input.watch.watchPath(targetPath, listener),
    })(),
  );

  const configHotReloadRuntime = createConfigHotReloadRuntime(
    createBuildConfigHotReloadRuntimeMainDepsHandler({
      getCurrentConfig: () => input.getResolvedConfig(),
      reloadConfigStrict: () => input.config.reloadConfigStrict(),
      watchConfigPath: (configPath, onChange) => watchConfigPathHandler(configPath, onChange),
      setTimeout: (callback, delayMs) => input.timers.setTimeout(callback, delayMs),
      clearTimeout: (timeout) => input.timers.clearTimeout(timeout),
      debounceMs: 250,
      onHotReloadApplied: createConfigHotReloadAppliedHandler(
        createBuildConfigHotReloadAppliedMainDepsHandler({
          setKeybindings: (keybindings) => {
            input.appState.keybindings = keybindings;
          },
          refreshGlobalAndOverlayShortcuts: () => {
            input.hotReload.refreshGlobalAndOverlayShortcuts();
          },
          setSecondarySubMode: (mode) => input.hotReload.setSecondarySubMode(mode),
          broadcastToOverlayWindows: (channel, payload) =>
            input.hotReload.broadcastToOverlayWindows(channel, payload),
          applyAnkiRuntimeConfigPatch: (patch) => {
            input.appState.ankiIntegration?.applyRuntimeConfigPatch(patch);
          },
        })(),
      ),
      onRestartRequired: (fields) =>
        notifyConfigHotReloadMessage(buildRestartRequiredConfigMessage(fields)),
      onInvalidConfig: notifyConfigHotReloadMessage,
      onValidationWarnings: (configPath, warnings: ConfigValidationWarning[]) => {
        input.notifications.showDesktopNotification('SubMiner', {
          body: buildConfigWarningNotificationBody(configPath, warnings),
        });
        if (input.platform === 'darwin') {
          input.notifications.showErrorBox(
            'SubMiner config validation warning',
            buildConfigWarningDialogDetails(configPath, warnings),
          );
        }
      },
    })(),
  );

  return {
    applyJellyfinMpvDefaults: (client) => applyJellyfinMpvDefaultsHandler(client),
    getDefaultSocketPath: () => getDefaultSocketPathHandler(),
    immersionMediaRuntime,
    configDerivedRuntime,
    subsyncRuntime,
    configHotReloadRuntime,
  };
}
