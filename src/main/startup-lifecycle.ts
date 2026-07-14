import { CliArgs, CliCommandSource } from '../cli/args';
import { createAppLifecycleDepsRuntime } from '../core/services';
import { startAppLifecycle } from '../core/services/app-lifecycle';
import type { AppLifecycleDepsRuntimeOptions } from '../core/services/app-lifecycle';
import { createAppLifecycleRuntimeDeps } from './app-lifecycle';

export interface AppLifecycleRuntimeRunnerParams {
  app: AppLifecycleDepsRuntimeOptions['app'];
  platform: NodeJS.Platform;
  shouldStartApp: (args: CliArgs) => boolean;
  parseArgs: (argv: string[]) => CliArgs;
  handleCliCommand: (nextArgs: CliArgs, source: CliCommandSource) => void;
  printHelp: () => void;
  logNoRunningInstance: () => void;
  startControlServer?: (handleArgv: (argv: string[]) => void) => (() => void) | void;
  onReady: () => Promise<void>;
  onWillQuitCleanup: () => void | Promise<void>;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
  shouldQuitOnWindowAllClosed: () => boolean;
}

export function createAppLifecycleRuntimeRunner(
  params: AppLifecycleRuntimeRunnerParams,
): (args: CliArgs) => void {
  return (args: CliArgs): void => {
    startAppLifecycle(
      args,
      createAppLifecycleDepsRuntime(
        createAppLifecycleRuntimeDeps({
          app: params.app,
          platform: params.platform,
          shouldStartApp: params.shouldStartApp,
          parseArgs: params.parseArgs,
          handleCliCommand: params.handleCliCommand,
          printHelp: params.printHelp,
          logNoRunningInstance: params.logNoRunningInstance,
          startControlServer: params.startControlServer,
          onReady: params.onReady,
          onWillQuitCleanup: params.onWillQuitCleanup,
          shouldRestoreWindowsOnActivate: params.shouldRestoreWindowsOnActivate,
          restoreWindowsOnActivate: params.restoreWindowsOnActivate,
          shouldQuitOnWindowAllClosed: params.shouldQuitOnWindowAllClosed,
        }),
      ),
    );
  };
}
