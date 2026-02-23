import { launchTexthookerOnly, runAppCommandWithInherit } from '../mpv.js';
import type { LauncherCommandContext } from './context.js';

export function runAppPassthroughCommand(context: LauncherCommandContext): boolean {
  const { args, appPath } = context;
  if (!args.appPassthrough || !appPath) {
    return false;
  }
  runAppCommandWithInherit(appPath, args.appArgs);
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
