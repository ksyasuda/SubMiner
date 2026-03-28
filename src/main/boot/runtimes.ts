import { createOpenFirstRunSetupWindowHandler } from '../runtime/first-run-setup-window';
import { createRunStatsCliCommandHandler } from '../runtime/stats-cli-command';
import { createYomitanProfilePolicy } from '../runtime/yomitan-profile-policy';
import {
  createBuildOpenAnilistSetupWindowMainDepsHandler,
  createMaybeFocusExistingAnilistSetupWindowHandler,
  createOpenAnilistSetupWindowHandler,
} from '../runtime/domains/anilist';
import {
  createTrayRuntimeHandlers,
  createYomitanExtensionRuntime,
  createYomitanSettingsRuntime,
} from '../runtime/domains/overlay';
import {
  composeAnilistSetupHandlers,
  composeAnilistTrackingHandlers,
  composeAppReadyRuntime,
  composeJellyfinRuntimeHandlers,
  composeMpvRuntimeHandlers,
  composeOverlayVisibilityRuntime,
  composeStatsStartupRuntime,
} from '../runtime/composers';

export interface MainBootRuntimesParams<TBrowserWindow, TMpvClient, TTokenizerDeps, TSubtitleData> {
  overlayVisibilityRuntimeDeps: Parameters<typeof composeOverlayVisibilityRuntime>[0];
  jellyfinRuntimeHandlerDeps: Parameters<typeof composeJellyfinRuntimeHandlers>[0];
  anilistSetupDeps: Parameters<typeof composeAnilistSetupHandlers>[0];
  buildOpenAnilistSetupWindowMainDeps: Parameters<
    typeof createBuildOpenAnilistSetupWindowMainDepsHandler
  >[0];
  anilistTrackingDeps: Parameters<typeof composeAnilistTrackingHandlers>[0];
  statsStartupRuntimeDeps: Parameters<typeof composeStatsStartupRuntime>[0];
  runStatsCliCommandDeps: Parameters<typeof createRunStatsCliCommandHandler>[0];
  appReadyRuntimeDeps: Parameters<typeof composeAppReadyRuntime>[0];
  mpvRuntimeDeps: any;
  trayRuntimeDeps: Parameters<typeof createTrayRuntimeHandlers>[0];
  yomitanProfilePolicyDeps: Parameters<typeof createYomitanProfilePolicy>[0];
  yomitanExtensionRuntimeDeps: Parameters<typeof createYomitanExtensionRuntime>[0];
  yomitanSettingsRuntimeDeps: Parameters<typeof createYomitanSettingsRuntime>[0];
  createOverlayRuntimeBootstrapHandlers: (params: {
    initializeOverlayRuntimeMainDeps: unknown;
    initializeOverlayRuntimeBootstrapDeps: unknown;
  }) => {
    initializeOverlayRuntime: () => void;
  };
  initializeOverlayRuntimeMainDeps: unknown;
  initializeOverlayRuntimeBootstrapDeps: unknown;
}

export function createMainBootRuntimes<
  TBrowserWindow,
  TMpvClient,
  TTokenizerDeps,
  TSubtitleData,
>(
  params: MainBootRuntimesParams<TBrowserWindow, TMpvClient, TTokenizerDeps, TSubtitleData>,
) {
  const overlayVisibilityComposer = composeOverlayVisibilityRuntime(
    params.overlayVisibilityRuntimeDeps,
  );
  const jellyfinRuntimeHandlers = composeJellyfinRuntimeHandlers(
    params.jellyfinRuntimeHandlerDeps,
  );
  const anilistSetupHandlers = composeAnilistSetupHandlers(params.anilistSetupDeps);
  const buildOpenAnilistSetupWindowMainDepsHandler =
    createBuildOpenAnilistSetupWindowMainDepsHandler(params.buildOpenAnilistSetupWindowMainDeps);
  const maybeFocusExistingAnilistSetupWindow =
    params.buildOpenAnilistSetupWindowMainDeps.maybeFocusExistingSetupWindow;
  const anilistTrackingHandlers = composeAnilistTrackingHandlers(params.anilistTrackingDeps);
  const statsStartupRuntime = composeStatsStartupRuntime(params.statsStartupRuntimeDeps);
  const runStatsCliCommand = createRunStatsCliCommandHandler(params.runStatsCliCommandDeps);
  const appReadyRuntime = composeAppReadyRuntime(params.appReadyRuntimeDeps);
  const mpvRuntimeHandlers = composeMpvRuntimeHandlers<any, any, any>(
    params.mpvRuntimeDeps as any,
  );
  const trayRuntimeHandlers = createTrayRuntimeHandlers(params.trayRuntimeDeps);
  const yomitanProfilePolicy = createYomitanProfilePolicy(params.yomitanProfilePolicyDeps);
  const yomitanExtensionRuntime = createYomitanExtensionRuntime(
    params.yomitanExtensionRuntimeDeps,
  );
  const yomitanSettingsRuntime = createYomitanSettingsRuntime(
    params.yomitanSettingsRuntimeDeps,
  );
  const overlayRuntimeBootstrapHandlers = params.createOverlayRuntimeBootstrapHandlers({
    initializeOverlayRuntimeMainDeps: params.initializeOverlayRuntimeMainDeps,
    initializeOverlayRuntimeBootstrapDeps: params.initializeOverlayRuntimeBootstrapDeps,
  });

  return {
    overlayVisibilityComposer,
    jellyfinRuntimeHandlers,
    anilistSetupHandlers,
    maybeFocusExistingAnilistSetupWindow,
    buildOpenAnilistSetupWindowMainDepsHandler,
    openAnilistSetupWindow: () =>
      createOpenAnilistSetupWindowHandler(buildOpenAnilistSetupWindowMainDepsHandler())(),
    anilistTrackingHandlers,
    statsStartupRuntime,
    runStatsCliCommand,
    appReadyRuntime,
    mpvRuntimeHandlers,
    trayRuntimeHandlers,
    yomitanProfilePolicy,
    yomitanExtensionRuntime,
    yomitanSettingsRuntime,
    initializeOverlayRuntime: overlayRuntimeBootstrapHandlers.initializeOverlayRuntime,
    openFirstRunSetupWindowHandler: createOpenFirstRunSetupWindowHandler,
  };
}

export const composeBootOverlayVisibilityRuntime = composeOverlayVisibilityRuntime;
export const composeBootJellyfinRuntimeHandlers = composeJellyfinRuntimeHandlers;
export const composeBootAnilistSetupHandlers = composeAnilistSetupHandlers;
export const composeBootAnilistTrackingHandlers = composeAnilistTrackingHandlers;
export const composeBootStatsStartupRuntime = composeStatsStartupRuntime;
export const createBootRunStatsCliCommandHandler = createRunStatsCliCommandHandler;
export const composeBootAppReadyRuntime = composeAppReadyRuntime;
export const composeBootMpvRuntimeHandlers = composeMpvRuntimeHandlers;
export const createBootTrayRuntimeHandlers = createTrayRuntimeHandlers;
export const createBootYomitanProfilePolicy = createYomitanProfilePolicy;
export const createBootYomitanExtensionRuntime = createYomitanExtensionRuntime;
export const createBootYomitanSettingsRuntime = createYomitanSettingsRuntime;
export const createBootMaybeFocusExistingAnilistSetupWindowHandler =
  createMaybeFocusExistingAnilistSetupWindowHandler;
export const createBootBuildOpenAnilistSetupWindowMainDepsHandler =
  createBuildOpenAnilistSetupWindowMainDepsHandler;
export const createBootOpenAnilistSetupWindowHandler = createOpenAnilistSetupWindowHandler;
