import {
  createDestroyTrayHandler,
  createEnsureOverlayVisibleFromTrayClickHandler,
  createEnsureTrayHandler,
} from './tray-lifecycle';
import {
  createBuildTrayMenuTemplateHandler,
  createResolveTrayIconPathHandler,
} from './tray-main-actions';

type ResolveTrayIconPathMainDeps = Parameters<typeof createResolveTrayIconPathHandler>[0];
type BuildTrayMenuTemplateMainDeps<TMenuItem> = Parameters<
  typeof createBuildTrayMenuTemplateHandler<TMenuItem>
>[0];
type EnsureTrayDeps = Parameters<typeof createEnsureTrayHandler>[0];
type DestroyTrayDeps = Parameters<typeof createDestroyTrayHandler>[0];

export function createTrayRuntimeHandlers<TMenuItem, TMenu>(deps: {
  resolveTrayIconPathDeps: ResolveTrayIconPathMainDeps;
  buildTrayMenuTemplateDeps: BuildTrayMenuTemplateMainDeps<TMenuItem>;
  ensureTrayDeps: Omit<
    EnsureTrayDeps,
    'buildTrayMenu' | 'resolveTrayIconPath' | 'ensureOverlayVisibleFromTrayClick'
  > &
    Parameters<typeof createEnsureOverlayVisibleFromTrayClickHandler>[0];
  destroyTrayDeps: DestroyTrayDeps;
  buildMenuFromTemplate: (template: TMenuItem[]) => TMenu;
}) {
  const resolveTrayIconPath = createResolveTrayIconPathHandler(deps.resolveTrayIconPathDeps);
  const buildTrayMenuTemplate = createBuildTrayMenuTemplateHandler(deps.buildTrayMenuTemplateDeps);
  const buildTrayMenu = () => deps.buildMenuFromTemplate(buildTrayMenuTemplate());

  const {
    initializeOverlayRuntime,
    isOverlayRuntimeInitialized,
    setVisibleOverlayVisible,
    ...ensureTrayDeps
  } = deps.ensureTrayDeps;
  const ensureTray = createEnsureTrayHandler({
    ...ensureTrayDeps,
    buildTrayMenu: () => buildTrayMenu(),
    resolveTrayIconPath: () => resolveTrayIconPath(),
    ensureOverlayVisibleFromTrayClick: createEnsureOverlayVisibleFromTrayClickHandler({
      initializeOverlayRuntime,
      isOverlayRuntimeInitialized,
      setVisibleOverlayVisible,
    }),
  });
  const destroyTray = createDestroyTrayHandler(deps.destroyTrayDeps);

  return {
    resolveTrayIconPath,
    buildTrayMenu,
    ensureTray,
    destroyTray,
  };
}
