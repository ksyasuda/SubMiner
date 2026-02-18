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
  onReady: () => Promise<void>;
  onWillQuitCleanup: () => void;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
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
          onReady: params.onReady,
          onWillQuitCleanup: params.onWillQuitCleanup,
          shouldRestoreWindowsOnActivate: params.shouldRestoreWindowsOnActivate,
          restoreWindowsOnActivate: params.restoreWindowsOnActivate,
        }),
      ),
    );
  };
}
