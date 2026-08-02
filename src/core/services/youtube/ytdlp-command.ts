import fs from 'node:fs';
import path from 'node:path';

const DEFAULT_YTDLP_COMMAND = 'yt-dlp';
const WINDOWS_YTDLP_COMMANDS = ['yt-dlp.cmd', 'yt-dlp.exe', 'yt-dlp'];

/**
 * yt-dlp expands `list=`/`index=` URL params into the whole playlist unless told not to, which
 * makes single-video extraction hang (e.g. a full Watch Later list) until our timeouts fire.
 * Every yt-dlp invocation targeting one video must include this.
 */
export const YTDLP_SINGLE_VIDEO_ARG = '--no-playlist';

function resolveFromPath(commandName: string): string | null {
  if (!process.env.PATH) {
    return null;
  }

  const searchPaths = process.env.PATH.split(path.delimiter);
  for (const searchPath of searchPaths) {
    const candidate = path.join(searchPath, commandName);
    try {
      fs.accessSync(candidate, fs.constants.X_OK);
      return candidate;
    } catch {
      continue;
    }
  }

  return null;
}

export function getYoutubeYtDlpCommand(): string {
  const explicitCommand = process.env.SUBMINER_YTDLP_BIN?.trim();
  if (explicitCommand) {
    return explicitCommand;
  }

  if (process.platform !== 'win32') {
    return DEFAULT_YTDLP_COMMAND;
  }

  for (const commandName of WINDOWS_YTDLP_COMMANDS) {
    const resolved = resolveFromPath(commandName);
    if (resolved) {
      return resolved;
    }
  }

  return DEFAULT_YTDLP_COMMAND;
}
