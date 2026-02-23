import type { Config } from '../../types';
import type { ConfiguredShortcuts } from '../../core/utils/shortcut-config';
import type { RegisterGlobalShortcutsServiceOptions } from '../../core/services/shortcut';

export function createBuildGetConfiguredShortcutsMainDepsHandler(deps: {
  getResolvedConfig: () => Config;
  defaultConfig: Config;
  resolveConfiguredShortcuts: (config: Config, defaultConfig: Config) => ConfiguredShortcuts;
}) {
  return () => ({
    getResolvedConfig: () => deps.getResolvedConfig(),
    defaultConfig: deps.defaultConfig,
    resolveConfiguredShortcuts: (config: Config, defaultConfig: Config) =>
      deps.resolveConfiguredShortcuts(config, defaultConfig),
  });
}

export function createBuildRegisterGlobalShortcutsMainDepsHandler(deps: {
  getConfiguredShortcuts: () => RegisterGlobalShortcutsServiceOptions['shortcuts'];
  registerGlobalShortcutsCore: (options: RegisterGlobalShortcutsServiceOptions) => void;
  toggleVisibleOverlay: () => void;
  toggleInvisibleOverlay: () => void;
  openYomitanSettings: () => void;
  isDev: boolean;
  getMainWindow: RegisterGlobalShortcutsServiceOptions['getMainWindow'];
}) {
  return () => ({
    getConfiguredShortcuts: () => deps.getConfiguredShortcuts(),
    registerGlobalShortcutsCore: (options: RegisterGlobalShortcutsServiceOptions) =>
      deps.registerGlobalShortcutsCore(options),
    onToggleVisibleOverlay: () => deps.toggleVisibleOverlay(),
    onToggleInvisibleOverlay: () => deps.toggleInvisibleOverlay(),
    onOpenYomitanSettings: () => deps.openYomitanSettings(),
    isDev: deps.isDev,
    getMainWindow: deps.getMainWindow,
  });
}

export function createBuildRefreshGlobalAndOverlayShortcutsMainDepsHandler(deps: {
  unregisterAllGlobalShortcuts: () => void;
  registerGlobalShortcuts: () => void;
  syncOverlayShortcuts: () => void;
}) {
  return () => ({
    unregisterAllGlobalShortcuts: () => deps.unregisterAllGlobalShortcuts(),
    registerGlobalShortcuts: () => deps.registerGlobalShortcuts(),
    syncOverlayShortcuts: () => deps.syncOverlayShortcuts(),
  });
}
