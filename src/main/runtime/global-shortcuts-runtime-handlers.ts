import type { ConfiguredShortcuts } from '../../core/utils/shortcut-config';
import {
  createGetConfiguredShortcutsHandler,
  createRefreshGlobalAndOverlayShortcutsHandler,
  createRegisterGlobalShortcutsHandler,
} from './global-shortcuts';
import {
  createBuildGetConfiguredShortcutsMainDepsHandler,
  createBuildRefreshGlobalAndOverlayShortcutsMainDepsHandler,
  createBuildRegisterGlobalShortcutsMainDepsHandler,
} from './global-shortcuts-main-deps';

type GetConfiguredShortcutsMainDeps = Parameters<
  typeof createBuildGetConfiguredShortcutsMainDepsHandler
>[0];
type RegisterGlobalShortcutsMainDeps = Parameters<
  typeof createBuildRegisterGlobalShortcutsMainDepsHandler
>[0];
type RefreshGlobalAndOverlayShortcutsMainDeps = Parameters<
  typeof createBuildRefreshGlobalAndOverlayShortcutsMainDepsHandler
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
  const getConfiguredShortcutsMainDeps = createBuildGetConfiguredShortcutsMainDepsHandler(
    deps.getConfiguredShortcutsMainDeps,
  )();
  const getConfiguredShortcutsHandler = createGetConfiguredShortcutsHandler(
    getConfiguredShortcutsMainDeps,
  );
  const getConfiguredShortcuts = () => getConfiguredShortcutsHandler();

  const registerGlobalShortcutsMainDeps = createBuildRegisterGlobalShortcutsMainDepsHandler(
    deps.buildRegisterGlobalShortcutsMainDeps(getConfiguredShortcuts),
  )();
  const registerGlobalShortcutsHandler = createRegisterGlobalShortcutsHandler(
    registerGlobalShortcutsMainDeps,
  );
  const registerGlobalShortcuts = () => registerGlobalShortcutsHandler();

  const refreshGlobalAndOverlayShortcutsMainDeps =
    createBuildRefreshGlobalAndOverlayShortcutsMainDepsHandler(
      deps.buildRefreshGlobalAndOverlayShortcutsMainDeps(registerGlobalShortcuts),
    )();
  const refreshGlobalAndOverlayShortcutsHandler = createRefreshGlobalAndOverlayShortcutsHandler(
    refreshGlobalAndOverlayShortcutsMainDeps,
  );
  const refreshGlobalAndOverlayShortcuts = () => refreshGlobalAndOverlayShortcutsHandler();

  return {
    getConfiguredShortcuts,
    registerGlobalShortcuts,
    refreshGlobalAndOverlayShortcuts,
  };
}
