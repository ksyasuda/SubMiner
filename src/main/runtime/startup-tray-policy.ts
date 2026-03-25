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
