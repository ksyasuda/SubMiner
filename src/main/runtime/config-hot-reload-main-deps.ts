import type {
  ConfigHotReloadDiff,
  ConfigHotReloadRuntimeDeps,
} from '../../core/services/config-hot-reload';
import type { ReloadConfigStrictResult } from '../../config';
import type {
  ConfigHotReloadPayload,
  ConfigValidationWarning,
  ResolvedConfig,
  SecondarySubMode,
} from '../../types';
import type { createConfigHotReloadMessageHandler } from './config-hot-reload-handlers';

type ConfigWatchListener = (eventType: string, filename: string | null) => void;

export function createWatchConfigPathHandler(deps: {
  fileExists: (path: string) => boolean;
  dirname: (path: string) => string;
  watchPath: (targetPath: string, listener: ConfigWatchListener) => { close: () => void };
}) {
  return (configPath: string, onChange: () => void): { close: () => void } => {
    const watchTarget = deps.fileExists(configPath) ? configPath : deps.dirname(configPath);
    const watcher = deps.watchPath(watchTarget, (_eventType, filename) => {
      if (watchTarget === configPath) {
        onChange();
        return;
      }

      const normalized =
        typeof filename === 'string' ? filename : filename ? String(filename) : undefined;
      if (!normalized || normalized === 'config.json' || normalized === 'config.jsonc') {
        onChange();
      }
    });
    return {
      close: () => watcher.close(),
    };
  };
}

type WatchConfigPathMainDeps = Parameters<typeof createWatchConfigPathHandler>[0];
type ConfigHotReloadMessageMainDeps = Parameters<typeof createConfigHotReloadMessageHandler>[0];

export function createBuildWatchConfigPathMainDepsHandler(deps: WatchConfigPathMainDeps) {
  return (): WatchConfigPathMainDeps => ({
    fileExists: (targetPath: string) => deps.fileExists(targetPath),
    dirname: (targetPath: string) => deps.dirname(targetPath),
    watchPath: (targetPath: string, listener: ConfigWatchListener) =>
      deps.watchPath(targetPath, listener),
  });
}

export function createBuildConfigHotReloadMessageMainDepsHandler(
  deps: ConfigHotReloadMessageMainDeps,
) {
  return (): ConfigHotReloadMessageMainDeps => ({
    showMpvOsd: (message: string) => deps.showMpvOsd(message),
    showDesktopNotification: (title: string, options: { body: string }) =>
      deps.showDesktopNotification(title, options),
  });
}

export function createBuildConfigHotReloadAppliedMainDepsHandler(deps: {
  setKeybindings: (keybindings: ConfigHotReloadPayload['keybindings']) => void;
  refreshGlobalAndOverlayShortcuts: () => void;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  applyAnkiRuntimeConfigPatch: (patch: { ai: ResolvedConfig['ankiConnect']['ai'] }) => void;
}) {
  return () => ({
    setKeybindings: (keybindings: ConfigHotReloadPayload['keybindings']) =>
      deps.setKeybindings(keybindings),
    refreshGlobalAndOverlayShortcuts: () => deps.refreshGlobalAndOverlayShortcuts(),
    setSecondarySubMode: (mode: SecondarySubMode) => deps.setSecondarySubMode(mode),
    broadcastToOverlayWindows: (channel: string, payload: unknown) =>
      deps.broadcastToOverlayWindows(channel, payload),
    applyAnkiRuntimeConfigPatch: (patch: { ai: ResolvedConfig['ankiConnect']['ai'] }) =>
      deps.applyAnkiRuntimeConfigPatch(patch),
  });
}

export function createBuildConfigHotReloadRuntimeMainDepsHandler(deps: {
  getCurrentConfig: () => ResolvedConfig;
  reloadConfigStrict: () => ReloadConfigStrictResult;
  watchConfigPath: ConfigHotReloadRuntimeDeps['watchConfigPath'];
  setTimeout: (callback: () => void, delayMs: number) => NodeJS.Timeout;
  clearTimeout: (timeout: NodeJS.Timeout) => void;
  debounceMs: number;
  onHotReloadApplied: (diff: ConfigHotReloadDiff, config: ResolvedConfig) => void;
  onRestartRequired: (fields: string[]) => void;
  onInvalidConfig: (message: string) => void;
  onValidationWarnings: (configPath: string, warnings: ConfigValidationWarning[]) => void;
}) {
  return (): ConfigHotReloadRuntimeDeps => ({
    getCurrentConfig: () => deps.getCurrentConfig(),
    reloadConfigStrict: () => deps.reloadConfigStrict(),
    watchConfigPath: (configPath, onChange) => deps.watchConfigPath(configPath, onChange),
    setTimeout: (callback: () => void, delayMs: number) => deps.setTimeout(callback, delayMs),
    clearTimeout: (timeout: NodeJS.Timeout) => deps.clearTimeout(timeout),
    debounceMs: deps.debounceMs,
    onHotReloadApplied: (diff, config) => deps.onHotReloadApplied(diff, config),
    onRestartRequired: (fields: string[]) => deps.onRestartRequired(fields),
    onInvalidConfig: (message: string) => deps.onInvalidConfig(message),
    onValidationWarnings: (configPath: string, warnings: ConfigValidationWarning[]) =>
      deps.onValidationWarnings(configPath, warnings),
  });
}
