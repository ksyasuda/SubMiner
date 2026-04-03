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

interface HeadlessStartupRuntimeSharedInput<TStartupState> {
  appLifecycleRuntimeRunnerMainDeps?: AppLifecycleDepsRuntimeOptions;
  appLifecycle?: HeadlessStartupAppLifecycleInput;
  bootstrap: HeadlessStartupBootstrapInput;
  createAppLifecycleRuntimeRunner?: (
    params: AppLifecycleDepsRuntimeOptions,
  ) => (args: CliArgs) => void;
  applyStartupState: (startupState: TStartupState) => void;
}

export interface HeadlessStartupRuntimeDefaultInput<TStartupState>
  extends HeadlessStartupRuntimeSharedInput<TStartupState> {
  createStartupBootstrapRuntimeDeps?: undefined;
  runStartupBootstrapRuntime: (deps: StartupBootstrapRuntimeDeps) => TStartupState;
}

export interface HeadlessStartupRuntimeCustomInput<TStartupState, TStartupBootstrapRuntimeDeps>
  extends HeadlessStartupRuntimeSharedInput<TStartupState> {
  createStartupBootstrapRuntimeDeps: (
    deps: StartupBootstrapRuntimeFactoryDeps,
  ) => TStartupBootstrapRuntimeDeps;
  runStartupBootstrapRuntime: (deps: TStartupBootstrapRuntimeDeps) => TStartupState;
}

export type HeadlessStartupRuntimeInput<
  TStartupState,
  TStartupBootstrapRuntimeDeps = StartupBootstrapRuntimeDeps,
> =
  | HeadlessStartupRuntimeDefaultInput<TStartupState>
  | HeadlessStartupRuntimeCustomInput<TStartupState, TStartupBootstrapRuntimeDeps>;

export interface HeadlessStartupRuntime<TStartupState> {
  appLifecycleRuntimeRunner: (args: CliArgs) => void;
  runAndApplyStartupState: () => TStartupState;
}

function resolveAppLifecycleRuntimeRunnerMainDeps(
  input: Pick<
    HeadlessStartupRuntimeSharedInput<unknown>,
    'appLifecycleRuntimeRunnerMainDeps' | 'appLifecycle'
  >,
) {
  const appLifecycleRuntimeRunnerMainDeps =
    input.appLifecycleRuntimeRunnerMainDeps ?? input.appLifecycle;

  if (!appLifecycleRuntimeRunnerMainDeps) {
    throw new Error('Headless startup runtime needs app lifecycle runtime runner deps');
  }

  return createBuildAppLifecycleRuntimeRunnerMainDepsHandler(appLifecycleRuntimeRunnerMainDeps)();
}

function buildHeadlessStartupHandlersDeps<TStartupState>(
  input: HeadlessStartupRuntimeSharedInput<TStartupState>,
) {
  const appLifecycleRuntimeRunnerMainDeps =
    resolveAppLifecycleRuntimeRunnerMainDeps(input);

  return {
    appLifecycleRuntimeRunnerMainDeps,
    createAppLifecycleRuntimeRunner:
      input.createAppLifecycleRuntimeRunner ??
      ((params: AppLifecycleDepsRuntimeOptions) => (args: CliArgs) =>
        startAppLifecycle(args, createAppLifecycleDepsRuntime(createAppLifecycleRuntimeDeps(params)))),
    buildStartupBootstrapMainDeps: (startAppLifecycle: (args: CliArgs) => void) => ({
      ...input.bootstrap,
      startAppLifecycle,
    }),
    applyStartupState: input.applyStartupState,
  };
}

export function createHeadlessStartupRuntime<TStartupState>(
  input: HeadlessStartupRuntimeDefaultInput<TStartupState>,
): HeadlessStartupRuntime<TStartupState>;
export function createHeadlessStartupRuntime<TStartupState, TStartupBootstrapRuntimeDeps>(
  input: HeadlessStartupRuntimeCustomInput<TStartupState, TStartupBootstrapRuntimeDeps>,
): HeadlessStartupRuntime<TStartupState>;
export function createHeadlessStartupRuntime<TStartupState, TStartupBootstrapRuntimeDeps>(
  input: HeadlessStartupRuntimeInput<TStartupState, TStartupBootstrapRuntimeDeps>,
): HeadlessStartupRuntime<TStartupState> {
  const baseDeps = buildHeadlessStartupHandlersDeps(input);
  const { appLifecycleRuntimeRunner, runAndApplyStartupState } =
    'createStartupBootstrapRuntimeDeps' in input && input.createStartupBootstrapRuntimeDeps
      ? composeHeadlessStartupHandlers<CliArgs, TStartupState, TStartupBootstrapRuntimeDeps>({
          startupRuntimeHandlersDeps: {
            ...baseDeps,
            createStartupBootstrapRuntimeDeps: (deps: StartupBootstrapRuntimeFactoryDeps) =>
              input.createStartupBootstrapRuntimeDeps(deps),
            runStartupBootstrapRuntime: input.runStartupBootstrapRuntime,
          },
        })
      : composeHeadlessStartupHandlers<CliArgs, TStartupState, StartupBootstrapRuntimeDeps>({
          startupRuntimeHandlersDeps: {
            ...baseDeps,
            createStartupBootstrapRuntimeDeps: (deps: StartupBootstrapRuntimeFactoryDeps) =>
              createStartupBootstrapRuntimeDeps(deps),
            runStartupBootstrapRuntime: input.runStartupBootstrapRuntime,
          },
        });

  return {
    appLifecycleRuntimeRunner,
    runAndApplyStartupState,
  };
}
