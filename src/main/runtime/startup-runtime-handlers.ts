import type { AppLifecycleRuntimeRunnerParams } from '../startup-lifecycle';
import type { StartupBootstrapRuntimeFactoryDeps } from '../startup';

// Bootstrap deps as main.ts supplies them; the config-generation exit callbacks are built here.
type StartupBootstrapMainDeps = Omit<
  StartupBootstrapRuntimeFactoryDeps,
  'onConfigGenerated' | 'onGenerateConfigError'
> & {
  setExitCode: (code: number) => void;
  quitApp: () => void;
  logGenerateConfigError: (message: string) => void;
};

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

  const { setExitCode, quitApp, logGenerateConfigError, ...bootstrapDeps } =
    deps.buildStartupBootstrapMainDeps(appLifecycleRuntimeRunner);
  const startupBootstrapRuntimeFactoryDeps: StartupBootstrapRuntimeFactoryDeps = {
    ...bootstrapDeps,
    onConfigGenerated: (exitCode) => {
      setExitCode(exitCode);
      quitApp();
    },
    onGenerateConfigError: (error) => {
      logGenerateConfigError(`Failed to generate config: ${error.message}`);
      setExitCode(1);
      quitApp();
    },
  };

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
