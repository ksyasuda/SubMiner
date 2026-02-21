import { createAppReadyRuntimeRunner } from '../../app-lifecycle';
import { createBuildAppReadyRuntimeMainDepsHandler } from '../app-ready-main-deps';
import {
  createBuildCriticalConfigErrorMainDepsHandler,
  createBuildReloadConfigMainDepsHandler,
} from '../startup-config-main-deps';
import { createCriticalConfigErrorHandler, createReloadConfigHandler } from '../startup-config';
import { createBuildImmersionTrackerStartupMainDepsHandler } from '../immersion-startup-main-deps';
import { createImmersionTrackerStartupHandler } from '../immersion-startup';

type ReloadConfigMainDeps = Parameters<typeof createBuildReloadConfigMainDepsHandler>[0];
type CriticalConfigErrorMainDeps = Parameters<
  typeof createBuildCriticalConfigErrorMainDepsHandler
>[0];
type AppReadyRuntimeMainDeps = Parameters<typeof createBuildAppReadyRuntimeMainDepsHandler>[0];

export type AppReadyComposerOptions = {
  reloadConfigMainDeps: ReloadConfigMainDeps;
  criticalConfigErrorMainDeps: CriticalConfigErrorMainDeps;
  appReadyRuntimeMainDeps: Omit<AppReadyRuntimeMainDeps, 'reloadConfig' | 'onCriticalConfigErrors'>;
  immersionTrackerStartupMainDeps: Parameters<
    typeof createBuildImmersionTrackerStartupMainDepsHandler
  >[0];
};

export type AppReadyComposerResult = {
  reloadConfig: ReturnType<typeof createReloadConfigHandler>;
  criticalConfigError: ReturnType<typeof createCriticalConfigErrorHandler>;
  appReadyRuntimeRunner: ReturnType<typeof createAppReadyRuntimeRunner>;
};

export function composeAppReadyRuntime(options: AppReadyComposerOptions): AppReadyComposerResult {
  const reloadConfig = createReloadConfigHandler(
    createBuildReloadConfigMainDepsHandler(options.reloadConfigMainDeps)(),
  );
  const criticalConfigError = createCriticalConfigErrorHandler(
    createBuildCriticalConfigErrorMainDepsHandler(options.criticalConfigErrorMainDeps)(),
  );

  const appReadyRuntimeRunner = createAppReadyRuntimeRunner(
    createBuildAppReadyRuntimeMainDepsHandler({
      ...options.appReadyRuntimeMainDeps,
      reloadConfig,
      createImmersionTracker: createImmersionTrackerStartupHandler(
        createBuildImmersionTrackerStartupMainDepsHandler(
          options.immersionTrackerStartupMainDeps,
        )(),
      ),
      onCriticalConfigErrors: criticalConfigError,
    })(),
  );

  return {
    reloadConfig,
    criticalConfigError,
    appReadyRuntimeRunner,
  };
}
