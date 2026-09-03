import {
  launchAppCommandDetached,
  launchAppBackgroundDetached,
  launchTexthookerOnly,
  runAppCommandWithInherit,
} from '../mpv.js';
import type { LauncherCommandContext } from './context.js';

type AppCommandDeps = {
  runAppCommandWithInherit: (appPath: string, appArgs: string[]) => void;
  launchSyncUiDetached: (
    appPath: string,
    logLevel: LauncherCommandContext['args']['logLevel'],
  ) => void;
  launchAppBackgroundDetached: (
    appPath: string,
    logLevel: LauncherCommandContext['args']['logLevel'],
  ) => void;
  launchAnimeBrowserDetached: (
    appPath: string,
    logLevel: LauncherCommandContext['args']['logLevel'],
  ) => void;
};

const defaultAppCommandDeps: AppCommandDeps = {
  runAppCommandWithInherit,
  launchSyncUiDetached: (appPath, logLevel) =>
    launchAppCommandDetached(appPath, ['--sync-window'], logLevel, 'sync-ui'),
  launchAppBackgroundDetached,
  launchAnimeBrowserDetached: (appPath, logLevel) =>
    launchAppCommandDetached(appPath, ['--anime'], logLevel, 'anime'),
};

export function runAppPassthroughCommand(
  context: LauncherCommandContext,
  deps: AppCommandDeps = defaultAppCommandDeps,
): boolean {
  const { args, appPath } = context;
  if (!appPath) {
    return false;
  }
  if (args.settings) {
    deps.runAppCommandWithInherit(appPath, ['--settings']);
    return true;
  }
  if (args.animeBrowser) {
    // Detached: the browser window is long-lived and owns the bridge process.
    deps.launchAnimeBrowserDetached(appPath, args.logLevel);
    return true;
  }
  if (args.syncUi) {
    deps.launchSyncUiDetached(appPath, args.logLevel);
    return true;
  }
  if (!args.appPassthrough) {
    return false;
  }
  if (args.appArgs.length === 0) {
    deps.launchAppBackgroundDetached(appPath, args.logLevel);
    return true;
  }
  deps.runAppCommandWithInherit(appPath, args.appArgs);
  return true;
}

export function runTexthookerCommand(context: LauncherCommandContext): boolean {
  const { args, appPath } = context;
  if (!args.texthookerOnly || !appPath) {
    return false;
  }
  launchTexthookerOnly(appPath, args);
  return true;
}
