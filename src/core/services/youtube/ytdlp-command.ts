const DEFAULT_YTDLP_COMMAND = 'yt-dlp';

export function getYoutubeYtDlpCommand(): string {
  return process.env.SUBMINER_YTDLP_BIN?.trim() || DEFAULT_YTDLP_COMMAND;
}
