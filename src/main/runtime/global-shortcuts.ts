import type { Config } from '../../types';
import type { ConfiguredShortcuts } from '../../core/utils/shortcut-config';
import type { RegisterGlobalShortcutsServiceOptions } from '../../core/services/shortcut';

export function createGetConfiguredShortcutsHandler(deps: {
  getResolvedConfig: () => Config;
  defaultConfig: Config;
  resolveConfiguredShortcuts: (
    config: Config,
    defaultConfig: Config,
  ) => ConfiguredShortcuts;
}) {
  return (): ConfiguredShortcuts =>
    deps.resolveConfiguredShortcuts(deps.getResolvedConfig(), deps.defaultConfig);
}

export function createRegisterGlobalShortcutsHandler(deps: {
  getConfiguredShortcuts: () => RegisterGlobalShortcutsServiceOptions['shortcuts'];
  registerGlobalShortcutsCore: (options: RegisterGlobalShortcutsServiceOptions) => void;
  onToggleVisibleOverlay: () => void;
  onOpenYomitanSettings: () => void;
  isDev: boolean;
  getMainWindow: RegisterGlobalShortcutsServiceOptions['getMainWindow'];
}) {
  return (): void => {
    deps.registerGlobalShortcutsCore({
      shortcuts: deps.getConfiguredShortcuts(),
      onToggleVisibleOverlay: deps.onToggleVisibleOverlay,
      onOpenYomitanSettings: deps.onOpenYomitanSettings,
      isDev: deps.isDev,
      getMainWindow: deps.getMainWindow,
    });
  };
}

export function createRefreshGlobalAndOverlayShortcutsHandler(deps: {
  unregisterAllGlobalShortcuts: () => void;
  registerGlobalShortcuts: () => void;
  syncOverlayShortcuts: () => void;
}) {
  return (): void => {
    deps.unregisterAllGlobalShortcuts();
    deps.registerGlobalShortcuts();
    deps.syncOverlayShortcuts();
  };
}
