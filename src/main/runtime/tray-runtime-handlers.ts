import { createDestroyTrayHandler, createEnsureTrayHandler } from './tray-lifecycle';
import {
  createBuildDestroyTrayMainDepsHandler,
  createBuildEnsureTrayMainDepsHandler,
} from './app-runtime-main-deps';
import {
  createBuildTrayMenuTemplateHandler,
  createResolveTrayIconPathHandler,
} from './tray-main-actions';
import {
  createBuildResolveTrayIconPathMainDepsHandler,
  createBuildTrayMenuTemplateMainDepsHandler,
} from './tray-main-deps';

type ResolveTrayIconPathMainDeps = Parameters<
  typeof createBuildResolveTrayIconPathMainDepsHandler
>[0];
type BuildTrayMenuTemplateMainDeps<TMenuItem> = Parameters<
  typeof createBuildTrayMenuTemplateMainDepsHandler<TMenuItem>
>[0];
type EnsureTrayMainDeps<TTrayMenu> = Parameters<
  typeof createBuildEnsureTrayMainDepsHandler<TrayLike, TTrayMenu, TrayIconLike>
>[0];
type TrayLike = NonNullable<ReturnType<Parameters<typeof createEnsureTrayHandler>[0]['getTray']>>;
type TrayIconLike = Parameters<Parameters<typeof createEnsureTrayHandler>[0]['createTray']>[0];
type DestroyTrayMainDeps = Parameters<typeof createBuildDestroyTrayMainDepsHandler<TrayLike>>[0];

export function createTrayRuntimeHandlers<TMenuItem, TMenu>(deps: {
  resolveTrayIconPathDeps: ResolveTrayIconPathMainDeps;
  buildTrayMenuTemplateDeps: BuildTrayMenuTemplateMainDeps<TMenuItem>;
  ensureTrayDeps: Omit<EnsureTrayMainDeps<TMenu>, 'buildTrayMenu' | 'resolveTrayIconPath'>;
  destroyTrayDeps: DestroyTrayMainDeps;
  buildMenuFromTemplate: (template: TMenuItem[]) => TMenu;
}) {
  const resolveTrayIconPath = createResolveTrayIconPathHandler(
    createBuildResolveTrayIconPathMainDepsHandler(deps.resolveTrayIconPathDeps)(),
  );
  const buildTrayMenuTemplate = createBuildTrayMenuTemplateHandler(
    createBuildTrayMenuTemplateMainDepsHandler(deps.buildTrayMenuTemplateDeps)(),
  );
  const buildTrayMenu = () => deps.buildMenuFromTemplate(buildTrayMenuTemplate());

  const ensureTray = createEnsureTrayHandler(
    createBuildEnsureTrayMainDepsHandler({
      ...deps.ensureTrayDeps,
      buildTrayMenu: () => buildTrayMenu(),
      resolveTrayIconPath: () => resolveTrayIconPath(),
    })(),
  );
  const destroyTray = createDestroyTrayHandler(
    createBuildDestroyTrayMainDepsHandler(deps.destroyTrayDeps)(),
  );

  return {
    resolveTrayIconPath,
    buildTrayMenu,
    ensureTray,
    destroyTray,
  };
}
