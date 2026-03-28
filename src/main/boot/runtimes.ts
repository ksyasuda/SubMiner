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
