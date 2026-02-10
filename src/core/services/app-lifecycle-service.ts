import { CliArgs, CliCommandSource } from "../../cli/args";

export interface AppLifecycleServiceDeps {
  shouldStartApp: (args: CliArgs) => boolean;
  parseArgs: (argv: string[]) => CliArgs;
  requestSingleInstanceLock: () => boolean;
  quitApp: () => void;
  onSecondInstance: (handler: (_event: unknown, argv: string[]) => void) => void;
  handleCliCommand: (args: CliArgs, source: CliCommandSource) => void;
  printHelp: () => void;
  logNoRunningInstance: () => void;
  whenReady: (handler: () => Promise<void>) => void;
  onWindowAllClosed: (handler: () => void) => void;
  onWillQuit: (handler: () => void) => void;
  onActivate: (handler: () => void) => void;
  isDarwinPlatform: () => boolean;
  onReady: () => Promise<void>;
  onWillQuitCleanup: () => void;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
}

export function startAppLifecycleService(
  initialArgs: CliArgs,
  deps: AppLifecycleServiceDeps,
): void {
  const gotTheLock = deps.requestSingleInstanceLock();
  if (!gotTheLock) {
    deps.quitApp();
    return;
  }

  deps.onSecondInstance((_event, argv) => {
    deps.handleCliCommand(deps.parseArgs(argv), "second-instance");
  });

  if (initialArgs.help && !deps.shouldStartApp(initialArgs)) {
    deps.printHelp();
    deps.quitApp();
    return;
  }

  if (!deps.shouldStartApp(initialArgs)) {
    if (initialArgs.stop && !initialArgs.start) {
      deps.quitApp();
    } else {
      deps.logNoRunningInstance();
      deps.quitApp();
    }
    return;
  }

  deps.whenReady(async () => {
    await deps.onReady();
  });

  deps.onWindowAllClosed(() => {
    if (!deps.isDarwinPlatform()) {
      deps.quitApp();
    }
  });

  deps.onWillQuit(() => {
    deps.onWillQuitCleanup();
  });

  deps.onActivate(() => {
    if (deps.shouldRestoreWindowsOnActivate()) {
      deps.restoreWindowsOnActivate();
    }
  });
}
