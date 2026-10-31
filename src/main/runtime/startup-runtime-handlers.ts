import type { AppLifecycleRuntimeRunnerParams } from '../startup-lifecycle';
import type { StartupBootstrapRuntimeFactoryDeps } from '../startup';
import { createBuildStartupBootstrapMainDepsHandler } from './startup-bootstrap-main-deps';

type StartupBootstrapMainDeps = Parameters<typeof createBuildStartupBootstrapMainDepsHandler>[0];

export function createStartupRuntimeHandlers<
  TCliArgs,
  TStartupState,
  TStartupBootstrapRuntimeDeps = unknown,
>(deps: {
  appLifecycleRuntimeRunnerMainDeps: AppLifecycleRuntimeRunnerParams;
  createAppLifecycleRuntimeRunner: (
    params: AppLifecycleRuntimeRunnerParams,
  ) => (args: TCliArgs) => void;
  buildStartupBootstrapMainDeps: (
    startAppLifecycle: (args: TCliArgs) => void,
  ) => StartupBootstrapMainDeps;
  createStartupBootstrapRuntimeDeps: (
    deps: StartupBootstrapRuntimeFactoryDeps,
  ) => TStartupBootstrapRuntimeDeps;
  runStartupBootstrapRuntime: (deps: TStartupBootstrapRuntimeDeps) => TStartupState;
  applyStartupState: (startupState: TStartupState) => void;
}) {
  const appLifecycleRuntimeRunner = deps.createAppLifecycleRuntimeRunner(
    deps.appLifecycleRuntimeRunnerMainDeps,
  );

  const startupBootstrapRuntimeFactoryDeps = createBuildStartupBootstrapMainDepsHandler(
    deps.buildStartupBootstrapMainDeps(appLifecycleRuntimeRunner),
  )();

  const runAndApplyStartupState = (): TStartupState => {
    const startupState = deps.runStartupBootstrapRuntime(
      deps.createStartupBootstrapRuntimeDeps(startupBootstrapRuntimeFactoryDeps),
    );
    deps.applyStartupState(startupState);
    return startupState;
  };

  return {
    appLifecycleRuntimeRunner,
    runAndApplyStartupState,
  };
}
