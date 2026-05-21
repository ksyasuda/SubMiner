import type { AppLifecycleRuntimeRunnerParams } from '../startup-lifecycle';

export function createBuildAppLifecycleRuntimeRunnerMainDepsHandler(
  deps: AppLifecycleRuntimeRunnerParams,
) {
  return (): AppLifecycleRuntimeRunnerParams => ({
    app: deps.app,
    platform: deps.platform,
    shouldStartApp: deps.shouldStartApp,
    parseArgs: deps.parseArgs,
    handleCliCommand: deps.handleCliCommand,
    printHelp: deps.printHelp,
    logNoRunningInstance: deps.logNoRunningInstance,
    startControlServer: deps.startControlServer,
    onReady: deps.onReady,
    onWillQuitCleanup: deps.onWillQuitCleanup,
    shouldRestoreWindowsOnActivate: deps.shouldRestoreWindowsOnActivate,
    restoreWindowsOnActivate: deps.restoreWindowsOnActivate,
    shouldQuitOnWindowAllClosed: deps.shouldQuitOnWindowAllClosed,
  });
}
