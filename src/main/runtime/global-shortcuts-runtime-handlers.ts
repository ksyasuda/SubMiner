import type { ConfiguredShortcuts } from '../../core/utils/shortcut-config';
import {
  createGetConfiguredShortcutsHandler,
  createRefreshGlobalAndOverlayShortcutsHandler,
  createRegisterGlobalShortcutsHandler,
} from './global-shortcuts';

type GetConfiguredShortcutsMainDeps = Parameters<typeof createGetConfiguredShortcutsHandler>[0];
type RegisterGlobalShortcutsMainDeps = Parameters<typeof createRegisterGlobalShortcutsHandler>[0];
type RefreshGlobalAndOverlayShortcutsMainDeps = Parameters<
  typeof createRefreshGlobalAndOverlayShortcutsHandler
>[0];

export function createGlobalShortcutsRuntimeHandlers(deps: {
  getConfiguredShortcutsMainDeps: GetConfiguredShortcutsMainDeps;
  buildRegisterGlobalShortcutsMainDeps: (
    getConfiguredShortcuts: () => ConfiguredShortcuts,
  ) => RegisterGlobalShortcutsMainDeps;
  buildRefreshGlobalAndOverlayShortcutsMainDeps: (
    registerGlobalShortcuts: () => void,
  ) => RefreshGlobalAndOverlayShortcutsMainDeps;
}) {
  const getConfiguredShortcutsHandler = createGetConfiguredShortcutsHandler(
    deps.getConfiguredShortcutsMainDeps,
  );
  const getConfiguredShortcuts = () => getConfiguredShortcutsHandler();

  const registerGlobalShortcutsHandler = createRegisterGlobalShortcutsHandler(
    deps.buildRegisterGlobalShortcutsMainDeps(getConfiguredShortcuts),
  );
  const registerGlobalShortcuts = () => registerGlobalShortcutsHandler();

  const refreshGlobalAndOverlayShortcutsHandler = createRefreshGlobalAndOverlayShortcutsHandler(
    deps.buildRefreshGlobalAndOverlayShortcutsMainDeps(registerGlobalShortcuts),
  );
  const refreshGlobalAndOverlayShortcuts = () => refreshGlobalAndOverlayShortcutsHandler();

  return {
    getConfiguredShortcuts,
    registerGlobalShortcuts,
    refreshGlobalAndOverlayShortcuts,
  };
}
