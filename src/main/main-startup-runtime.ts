import type { AppReadyRuntimeInput, AppReadyRuntime } from './app-ready-runtime';
import type { CliStartupRuntimeInput, CliStartupRuntime } from './cli-startup-runtime';
import type {
  HeadlessStartupRuntimeInput,
  HeadlessStartupRuntime,
} from './headless-startup-runtime';
import { createAppReadyRuntime } from './app-ready-runtime';
import { createCliStartupRuntime } from './cli-startup-runtime';
import { createHeadlessStartupRuntime } from './headless-startup-runtime';

export interface MainStartupRuntimeInput<TStartupState> {
  appReady: AppReadyRuntimeInput;
  cli: CliStartupRuntimeInput;
  headless: HeadlessStartupRuntimeInput<TStartupState>;
}

export interface MainStartupRuntime<TStartupState> {
  appReady: AppReadyRuntime;
  cliStartup: CliStartupRuntime;
  headlessStartup: HeadlessStartupRuntime<TStartupState>;
  handleCliCommand: CliStartupRuntime['handleCliCommand'];
  handleInitialArgs: CliStartupRuntime['handleInitialArgs'];
  appLifecycleRuntimeRunner: HeadlessStartupRuntime<TStartupState>['appLifecycleRuntimeRunner'];
  runAndApplyStartupState: HeadlessStartupRuntime<TStartupState>['runAndApplyStartupState'];
}

export function createMainStartupRuntime<TStartupState>(
  input: MainStartupRuntimeInput<TStartupState>,
): MainStartupRuntime<TStartupState> {
  const appReady = createAppReadyRuntime(input.appReady);
  const cliStartup = createCliStartupRuntime(input.cli);
  const headlessStartup = createHeadlessStartupRuntime<TStartupState>(input.headless);

  return {
    appReady,
    cliStartup,
    headlessStartup,
    handleCliCommand: cliStartup.handleCliCommand,
    handleInitialArgs: cliStartup.handleInitialArgs,
    appLifecycleRuntimeRunner: headlessStartup.appLifecycleRuntimeRunner,
    runAndApplyStartupState: headlessStartup.runAndApplyStartupState,
  };
}
