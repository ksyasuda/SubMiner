import {
  createAppReadyRuntimeRunner,
  type AppReadyRuntimeDepsFactoryInput,
} from '../../app-lifecycle';
import { createCriticalConfigErrorHandler, createReloadConfigHandler } from '../startup-config';
import { createBuildImmersionTrackerStartupMainDepsHandler } from '../immersion-startup-main-deps';
import { createImmersionTrackerStartupHandler } from '../immersion-startup';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type ReloadConfigMainDeps = Parameters<typeof createReloadConfigHandler>[0];
type CriticalConfigErrorMainDeps = Parameters<typeof createCriticalConfigErrorHandler>[0];

export type AppReadyComposerOptions = ComposerInputs<{
  reloadConfigMainDeps: ReloadConfigMainDeps;
  criticalConfigErrorMainDeps: CriticalConfigErrorMainDeps;
  appReadyRuntimeMainDeps: Omit<
    AppReadyRuntimeDepsFactoryInput,
    'reloadConfig' | 'onCriticalConfigErrors'
  >;
  immersionTrackerStartupMainDeps: Parameters<
    typeof createBuildImmersionTrackerStartupMainDepsHandler
  >[0];
}>;

export type AppReadyComposerResult = ComposerOutputs<{
  reloadConfig: ReturnType<typeof createReloadConfigHandler>;
  criticalConfigError: ReturnType<typeof createCriticalConfigErrorHandler>;
  appReadyRuntimeRunner: ReturnType<typeof createAppReadyRuntimeRunner>;
}>;

export function composeAppReadyRuntime(options: AppReadyComposerOptions): AppReadyComposerResult {
  const reloadConfig = createReloadConfigHandler(options.reloadConfigMainDeps);
  const criticalConfigError = createCriticalConfigErrorHandler(options.criticalConfigErrorMainDeps);

  const appReadyRuntimeRunner = createAppReadyRuntimeRunner({
    ...options.appReadyRuntimeMainDeps,
    reloadConfig,
    createImmersionTracker:
      options.appReadyRuntimeMainDeps.createImmersionTracker ??
      createImmersionTrackerStartupHandler(
        createBuildImmersionTrackerStartupMainDepsHandler(
          options.immersionTrackerStartupMainDeps,
        )(),
      ),
    onCriticalConfigErrors: criticalConfigError,
  });

  return {
    reloadConfig,
    criticalConfigError,
    appReadyRuntimeRunner,
  };
}
