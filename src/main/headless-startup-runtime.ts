import type { CliArgs } from '../cli/args';
import type { LogLevelSource } from '../logger';
import type { ResolvedConfig } from '../types';
import type { StartupBootstrapRuntimeDeps } from '../core/services/startup';
import { createAppLifecycleDepsRuntime, startAppLifecycle } from '../core/services/app-lifecycle';
import type { AppLifecycleDepsRuntimeOptions } from '../core/services/app-lifecycle';
import type { AppLifecycleRuntimeRunnerParams } from './startup-lifecycle';
import type { StartupBootstrapRuntimeFactoryDeps } from './startup';
import { createStartupBootstrapRuntimeDeps } from './startup';
import { composeHeadlessStartupHandlers } from './runtime/composers/headless-startup-composer';
import { createAppLifecycleRuntimeDeps } from './app-lifecycle';
import { createBuildAppLifecycleRuntimeRunnerMainDepsHandler } from './runtime/startup-lifecycle-main-deps';

export interface HeadlessStartupBootstrapInput {
  argv: string[];
  parseArgs: (argv: string[]) => CliArgs;
  setLogLevel: (level: string, source: LogLevelSource) => void;
  forceX11Backend: (args: CliArgs) => void;
  enforceUnsupportedWaylandMode: (args: CliArgs) => void;
  shouldStartApp: (args: CliArgs) => boolean;
  getDefaultSocketPath: () => string;
  defaultTexthookerPort: number;
  configDir: string;
  defaultConfig: ResolvedConfig;
  generateConfigTemplate: (config: ResolvedConfig) => string;
  generateDefaultConfigFile: (
    args: CliArgs,
    options: {
      configDir: string;
      defaultConfig: unknown;
      generateTemplate: (config: unknown) => string;
    },
  ) => Promise<number>;
  setExitCode: (code: number) => void;
  quitApp: () => void;
  logGenerateConfigError: (message: string) => void;
  startAppLifecycle: (args: CliArgs) => void;
}

export type HeadlessStartupAppLifecycleInput = AppLifecycleRuntimeRunnerParams;

export interface HeadlessStartupRuntimeInput<
  TStartupState,
  TStartupBootstrapRuntimeDeps = StartupBootstrapRuntimeDeps,
> {
  appLifecycleRuntimeRunnerMainDeps?: AppLifecycleDepsRuntimeOptions;
  appLifecycle?: HeadlessStartupAppLifecycleInput;
  bootstrap: HeadlessStartupBootstrapInput;
  createAppLifecycleRuntimeRunner?: (
    params: AppLifecycleDepsRuntimeOptions,
  ) => (args: CliArgs) => void;
  createStartupBootstrapRuntimeDeps?: (
    deps: StartupBootstrapRuntimeFactoryDeps,
  ) => TStartupBootstrapRuntimeDeps;
  runStartupBootstrapRuntime: (deps: TStartupBootstrapRuntimeDeps) => TStartupState;
  applyStartupState: (startupState: TStartupState) => void;
}

export interface HeadlessStartupRuntime<TStartupState> {
  appLifecycleRuntimeRunner: (args: CliArgs) => void;
  runAndApplyStartupState: () => TStartupState;
}

export function createHeadlessStartupRuntime<
  TStartupState,
  TStartupBootstrapRuntimeDeps = StartupBootstrapRuntimeDeps,
>(
  input: HeadlessStartupRuntimeInput<TStartupState, TStartupBootstrapRuntimeDeps>,
): HeadlessStartupRuntime<TStartupState> {
  const appLifecycleRuntimeRunnerMainDeps =
    input.appLifecycleRuntimeRunnerMainDeps ?? input.appLifecycle;

  if (!appLifecycleRuntimeRunnerMainDeps) {
    throw new Error('Headless startup runtime needs app lifecycle runtime runner deps');
  }

  const { appLifecycleRuntimeRunner, runAndApplyStartupState } = composeHeadlessStartupHandlers({
    startupRuntimeHandlersDeps: {
      appLifecycleRuntimeRunnerMainDeps: createBuildAppLifecycleRuntimeRunnerMainDepsHandler(
        appLifecycleRuntimeRunnerMainDeps,
      )(),
      createAppLifecycleRuntimeRunner:
        input.createAppLifecycleRuntimeRunner ??
        ((params: AppLifecycleDepsRuntimeOptions) => (args: CliArgs) =>
          startAppLifecycle(
            args,
            createAppLifecycleDepsRuntime(createAppLifecycleRuntimeDeps(params)),
          )),
      buildStartupBootstrapMainDeps: (startAppLifecycle: (args: CliArgs) => void) => ({
        ...input.bootstrap,
        startAppLifecycle,
      }),
      createStartupBootstrapRuntimeDeps: (deps: StartupBootstrapRuntimeFactoryDeps) =>
        input.createStartupBootstrapRuntimeDeps
          ? input.createStartupBootstrapRuntimeDeps(deps)
          : (createStartupBootstrapRuntimeDeps(deps) as unknown as TStartupBootstrapRuntimeDeps),
      runStartupBootstrapRuntime: input.runStartupBootstrapRuntime,
      applyStartupState: input.applyStartupState,
    },
  });

  return {
    appLifecycleRuntimeRunner,
    runAndApplyStartupState,
  };
}
