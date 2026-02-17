import { CliArgs, CliCommandSource } from "../../cli/args";
import { createLogger } from "../../logger";

const logger = createLogger("main:app-lifecycle");

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

interface AppLike {
  requestSingleInstanceLock: () => boolean;
  quit: () => void;
  on: (...args: any[]) => unknown;
  whenReady: () => Promise<void>;
}

export interface AppLifecycleDepsRuntimeOptions {
  app: AppLike;
  platform: NodeJS.Platform;
  shouldStartApp: (args: CliArgs) => boolean;
  parseArgs: (argv: string[]) => CliArgs;
  handleCliCommand: (args: CliArgs, source: CliCommandSource) => void;
  printHelp: () => void;
  logNoRunningInstance: () => void;
  onReady: () => Promise<void>;
  onWillQuitCleanup: () => void;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
}

export function createAppLifecycleDepsRuntime(
  options: AppLifecycleDepsRuntimeOptions,
): AppLifecycleServiceDeps {
  return {
    shouldStartApp: options.shouldStartApp,
    parseArgs: options.parseArgs,
    requestSingleInstanceLock: () => options.app.requestSingleInstanceLock(),
    quitApp: () => options.app.quit(),
    onSecondInstance: (handler) => {
      options.app.on("second-instance", handler as (...args: unknown[]) => void);
    },
    handleCliCommand: options.handleCliCommand,
    printHelp: options.printHelp,
    logNoRunningInstance: options.logNoRunningInstance,
    whenReady: (handler) => {
      options.app.whenReady().then(handler).catch((error) => {
        logger.error("App ready handler failed:", error);
      });
    },
    onWindowAllClosed: (handler) => {
      options.app.on("window-all-closed", handler as (...args: unknown[]) => void);
    },
    onWillQuit: (handler) => {
      options.app.on("will-quit", handler as (...args: unknown[]) => void);
    },
    onActivate: (handler) => {
      options.app.on("activate", handler as (...args: unknown[]) => void);
    },
    isDarwinPlatform: () => options.platform === "darwin",
    onReady: options.onReady,
    onWillQuitCleanup: options.onWillQuitCleanup,
    shouldRestoreWindowsOnActivate: options.shouldRestoreWindowsOnActivate,
    restoreWindowsOnActivate: options.restoreWindowsOnActivate,
  };
}

export function startAppLifecycle(
  initialArgs: CliArgs,
  deps: AppLifecycleServiceDeps,
): void {
  const gotTheLock = deps.requestSingleInstanceLock();
  if (!gotTheLock) {
    deps.quitApp();
    return;
  }

  deps.onSecondInstance((_event, argv) => {
    try {
      deps.handleCliCommand(deps.parseArgs(argv), "second-instance");
    } catch (error) {
      logger.error("Failed to handle second-instance CLI command:", error);
    }
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
