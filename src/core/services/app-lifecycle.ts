import { CliArgs, CliCommandSource } from '../../cli/args';
import { createLogger } from '../../logger';

const logger = createLogger('main:app-lifecycle');

export interface AppLifecycleServiceDeps {
  shouldStartApp: (args: CliArgs) => boolean;
  parseArgs: (argv: string[]) => CliArgs;
  requestSingleInstanceLock: () => boolean;
  quitApp: () => void;
  exitApp: (code: number) => void;
  onSecondInstance: (handler: (_event: unknown, argv: string[]) => void) => void;
  handleCliCommand: (args: CliArgs, source: CliCommandSource) => void;
  printHelp: () => void;
  logNoRunningInstance: () => void;
  startControlServer?: (handleArgv: (argv: string[]) => void) => (() => void) | void;
  whenReady: (handler: () => Promise<void>) => void;
  onWindowAllClosed: (handler: () => void) => void;
  onWillQuit: (handler: (event: { preventDefault(): void }) => void) => void;
  onActivate: (handler: () => void) => void;
  isDarwinPlatform: () => boolean;
  onReady: () => Promise<void>;
  onWillQuitCleanup: () => void | Promise<void>;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
  shouldQuitOnWindowAllClosed: () => boolean;
}

interface AppLike {
  requestSingleInstanceLock: () => boolean;
  quit: () => void;
  exit?: (exitCode?: number) => void;
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
  startControlServer?: (handleArgv: (argv: string[]) => void) => (() => void) | void;
  onReady: () => Promise<void>;
  onWillQuitCleanup: () => void | Promise<void>;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
  shouldQuitOnWindowAllClosed: () => boolean;
}

export function createAppLifecycleDepsRuntime(
  options: AppLifecycleDepsRuntimeOptions,
): AppLifecycleServiceDeps {
  return {
    shouldStartApp: options.shouldStartApp,
    parseArgs: options.parseArgs,
    requestSingleInstanceLock: () => options.app.requestSingleInstanceLock(),
    quitApp: () => options.app.quit(),
    exitApp: (code) => {
      if (options.app.exit) {
        options.app.exit(code);
        return;
      }
      process.exitCode = code;
      options.app.quit();
    },
    onSecondInstance: (handler) => {
      options.app.on('second-instance', handler as (...args: unknown[]) => void);
    },
    handleCliCommand: options.handleCliCommand,
    printHelp: options.printHelp,
    logNoRunningInstance: options.logNoRunningInstance,
    startControlServer: options.startControlServer,
    whenReady: (handler) => {
      options.app
        .whenReady()
        .then(handler)
        .catch((error) => {
          logger.error('App ready handler failed:', error);
        });
    },
    onWindowAllClosed: (handler) => {
      options.app.on('window-all-closed', handler as (...args: unknown[]) => void);
    },
    onWillQuit: (handler) => {
      options.app.on('will-quit', handler as (...args: unknown[]) => void);
    },
    onActivate: (handler) => {
      options.app.on('activate', handler as (...args: unknown[]) => void);
    },
    isDarwinPlatform: () => options.platform === 'darwin',
    onReady: options.onReady,
    onWillQuitCleanup: options.onWillQuitCleanup,
    shouldRestoreWindowsOnActivate: options.shouldRestoreWindowsOnActivate,
    restoreWindowsOnActivate: options.restoreWindowsOnActivate,
    shouldQuitOnWindowAllClosed: options.shouldQuitOnWindowAllClosed,
  };
}

export function startAppLifecycle(initialArgs: CliArgs, deps: AppLifecycleServiceDeps): void {
  if (initialArgs.help && !deps.shouldStartApp(initialArgs)) {
    deps.printHelp();
    deps.quitApp();
    return;
  }

  const gotTheLock = deps.requestSingleInstanceLock();
  if (initialArgs.appPing) {
    deps.exitApp(gotTheLock ? 1 : 0);
    return;
  }

  if (!gotTheLock) {
    deps.quitApp();
    return;
  }

  let appReadyRuntimeComplete = false;
  const pendingSecondInstanceCommands: CliArgs[] = [];
  let stopControlServer: (() => void) | null = null;
  const handleSecondInstanceCommand = (args: CliArgs): void => {
    try {
      deps.handleCliCommand(args, 'second-instance');
    } catch (error) {
      logger.error('Failed to handle second-instance CLI command:', error);
    }
  };

  const flushPendingSecondInstanceCommands = (): void => {
    while (pendingSecondInstanceCommands.length > 0) {
      const nextArgs = pendingSecondInstanceCommands.shift();
      if (nextArgs) {
        handleSecondInstanceCommand(nextArgs);
      }
    }
  };

  const dispatchSecondInstanceArgv = (argv: string[]): void => {
    try {
      const nextArgs = deps.parseArgs(argv);
      if (!appReadyRuntimeComplete) {
        pendingSecondInstanceCommands.push(nextArgs);
        return;
      }

      handleSecondInstanceCommand(nextArgs);
    } catch (error) {
      logger.error('Failed to handle second-instance CLI command:', error);
    }
  };

  deps.onSecondInstance((_event, argv) => {
    dispatchSecondInstanceArgv(argv);
  });

  if (!deps.shouldStartApp(initialArgs)) {
    if (initialArgs.stop && !initialArgs.start) {
      deps.quitApp();
    } else {
      deps.logNoRunningInstance();
      deps.quitApp();
    }
    return;
  }

  try {
    stopControlServer = deps.startControlServer?.(dispatchSecondInstanceArgv) ?? null;
  } catch (error) {
    logger.error('Failed to start app control socket:', error);
  }

  deps.whenReady(async () => {
    try {
      await deps.onReady();
    } finally {
      appReadyRuntimeComplete = true;
      flushPendingSecondInstanceCommands();
    }
  });

  deps.onWindowAllClosed(() => {
    if (
      deps.shouldQuitOnWindowAllClosed() &&
      (!deps.isDarwinPlatform() ||
        initialArgs.settings ||
        initialArgs.setup ||
        initialArgs.syncWindow)
    ) {
      deps.quitApp();
    }
  });

  let quitCleanupPending = false;
  let quitCleanupComplete = false;
  deps.onWillQuit((event) => {
    if (quitCleanupComplete) return;
    stopControlServer?.();
    stopControlServer = null;
    if (quitCleanupPending) {
      event.preventDefault();
      return;
    }
    let cleanup: void | Promise<void>;
    try {
      cleanup = deps.onWillQuitCleanup();
    } catch (error) {
      logger.error('App quit cleanup failed:', error);
      return;
    }
    if (!(cleanup instanceof Promise)) return;
    quitCleanupPending = true;
    event.preventDefault();
    void cleanup
      .catch((error) => {
        logger.error('App quit cleanup failed:', error);
      })
      .finally(() => {
        quitCleanupPending = false;
        quitCleanupComplete = true;
        deps.quitApp();
      });
  });

  deps.onActivate(() => {
    if (deps.shouldRestoreWindowsOnActivate()) {
      deps.restoreWindowsOnActivate();
    }
  });
}
