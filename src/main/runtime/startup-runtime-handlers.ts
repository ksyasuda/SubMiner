import { createBuildStartupBootstrapRuntimeFactoryDepsHandler } from './startup-bootstrap-deps-builder';
import type { AppLifecycleRuntimeRunnerParams } from '../startup-lifecycle';
import { createBuildStartupBootstrapMainDepsHandler } from './startup-bootstrap-main-deps';

type StartupBootstrapMainDeps = Parameters<typeof createBuildStartupBootstrapMainDepsHandler>[0];
type StartupBootstrapRuntimeFactoryDeps = Parameters<
  typeof createBuildStartupBootstrapRuntimeFactoryDepsHandler
>[0];

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

  const startupBootstrapMainDeps = createBuildStartupBootstrapMainDepsHandler(
    deps.buildStartupBootstrapMainDeps(appLifecycleRuntimeRunner),
  )();
  const startupBootstrapRuntimeFactoryDeps =
    createBuildStartupBootstrapRuntimeFactoryDepsHandler(startupBootstrapMainDeps)();

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
