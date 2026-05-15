import type { CliArgs } from '../../cli/args';

export function shouldEnsureTrayOnStartupForInitialArgs(
  platform: NodeJS.Platform,
  initialArgs: CliArgs | null,
): boolean {
  if (platform !== 'win32') {
    return false;
  }
  if (initialArgs?.youtubePlay) {
    return false;
  }
  return true;
}

export function shouldQuitOnWindowAllClosedForTrayState(options: {
  backgroundMode: boolean;
  hasTray: boolean;
}): boolean {
  if (options.backgroundMode) return false;
  if (options.hasTray) return false;
  return true;
}
