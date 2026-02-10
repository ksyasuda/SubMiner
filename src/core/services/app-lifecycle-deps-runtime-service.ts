import { CliArgs, CliCommandSource } from "../../cli/args";
import { AppLifecycleServiceDeps } from "./app-lifecycle-service";

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

export function createAppLifecycleDepsRuntimeService(
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
      options.app.whenReady().then(handler);
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
