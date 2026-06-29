import type { CliArgs, CliCommandSource } from '../../cli/args';

type LaunchResult = {
  ok: boolean;
  mpvPath?: string;
};

export type YoutubePlaybackRuntimeDeps = {
  platform: NodeJS.Platform;
  directPlaybackFormat: string;
  mpvYtdlFormat: string;
  autoLaunchTimeoutMs: number;
  connectTimeoutMs: number;
  getSocketPath: () => string;
  getMpvConnected: () => boolean;
  invalidatePendingAutoplayReadyFallbacks: () => void;
  setAppOwnedFlowInFlight: (next: boolean) => void;
  ensureYoutubePlaybackRuntimeReady: () => Promise<void>;
  resolveYoutubePlaybackUrl: (url: string, format: string) => Promise<string>;
  launchWindowsMpv: (playbackUrl: string, args: string[]) => Promise<LaunchResult>;
  waitForYoutubeMpvConnected: (timeoutMs: number) => Promise<boolean>;
  prepareYoutubePlaybackInMpv: (request: { url: string }) => Promise<boolean>;
  startYoutubeMediaCache?: (url: string) => void | Promise<void>;
  runYoutubePlaybackFlow: (request: {
    url: string;
    mode: NonNullable<CliArgs['youtubeMode']>;
  }) => Promise<void>;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  clearScheduled: (timer: ReturnType<typeof setTimeout>) => void;
};

export function createYoutubePlaybackRuntime(deps: YoutubePlaybackRuntimeDeps) {
  let quitOnDisconnectArmed = false;
  let quitOnDisconnectArmTimer: ReturnType<typeof setTimeout> | null = null;
  let playbackFlowGeneration = 0;

  const clearYoutubePlayQuitOnDisconnectArmTimer = (): void => {
    if (quitOnDisconnectArmTimer) {
      deps.clearScheduled(quitOnDisconnectArmTimer);
      quitOnDisconnectArmTimer = null;
    }
  };

  const runYoutubePlaybackFlow = async (request: {
    url: string;
    mode: NonNullable<CliArgs['youtubeMode']>;
    source: CliCommandSource;
  }): Promise<void> => {
    const flowGeneration = ++playbackFlowGeneration;
    deps.invalidatePendingAutoplayReadyFallbacks();
    deps.setAppOwnedFlowInFlight(true);
    let flowCompleted = false;

    try {
      clearYoutubePlayQuitOnDisconnectArmTimer();
      quitOnDisconnectArmed = false;
      await deps.ensureYoutubePlaybackRuntimeReady();

      let playbackUrl = request.url;
      let launchedWindowsMpv = false;
      if (deps.platform === 'win32') {
        try {
          playbackUrl = await deps.resolveYoutubePlaybackUrl(
            request.url,
            deps.directPlaybackFormat,
          );
          deps.logInfo('Resolved direct YouTube playback URL for Windows MPV startup.');
        } catch (error) {
          deps.logWarn(
            `Failed to resolve direct YouTube playback URL; falling back to page URL: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        }
      }

      if (deps.platform === 'win32' && !deps.getMpvConnected()) {
        const socketPath = deps.getSocketPath();
        const launchResult = await deps.launchWindowsMpv(playbackUrl, [
          '--pause=yes',
          '--ytdl=yes',
          `--ytdl-format=${deps.mpvYtdlFormat}`,
          '--sub-auto=no',
          '--sub-file-paths=.;subs;subtitles',
          '--sid=auto',
          '--secondary-sid=auto',
          '--secondary-sub-visibility=no',
          '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
          '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
          `--input-ipc-server=${socketPath}`,
        ]);
        launchedWindowsMpv = launchResult.ok;
        if (launchResult.ok && launchResult.mpvPath) {
          deps.logInfo(
            `Bootstrapping Windows mpv for YouTube playback via ${launchResult.mpvPath}`,
          );
        }
        if (!launchResult.ok) {
          deps.logWarn('Unable to bootstrap Windows mpv for YouTube playback.');
        }
      }

      const connected = await deps.waitForYoutubeMpvConnected(
        launchedWindowsMpv ? deps.autoLaunchTimeoutMs : deps.connectTimeoutMs,
      );
      if (!connected) {
        throw new Error(
          launchedWindowsMpv
            ? 'MPV not connected after auto-launch. Ensure mpv is installed and can open the requested YouTube URL.'
            : 'MPV not connected. Start mpv with the SubMiner profile or retry after mpv finishes starting.',
        );
      }

      if (request.source === 'initial') {
        quitOnDisconnectArmTimer = deps.schedule(() => {
          if (playbackFlowGeneration !== flowGeneration) {
            return;
          }
          quitOnDisconnectArmed = true;
          quitOnDisconnectArmTimer = null;
        }, 3000);
      }

      const mediaReady = await deps.prepareYoutubePlaybackInMpv({ url: playbackUrl });
      if (!mediaReady) {
        throw new Error('Timed out waiting for mpv to load the requested YouTube URL.');
      }
      if (deps.startYoutubeMediaCache) {
        void new Promise<void>((resolve) => {
          resolve(deps.startYoutubeMediaCache?.(request.url));
        }).catch((error) => {
          deps.logWarn(
            `Failed to start YouTube media cache: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        });
      }

      await deps.runYoutubePlaybackFlow({
        url: request.url,
        mode: request.mode,
      });
      flowCompleted = true;
      deps.logInfo(`YouTube playback flow completed from ${request.source}.`);
    } finally {
      if (playbackFlowGeneration === flowGeneration) {
        if (!flowCompleted) {
          clearYoutubePlayQuitOnDisconnectArmTimer();
          quitOnDisconnectArmed = false;
        }
        deps.setAppOwnedFlowInFlight(false);
      }
    }
  };

  return {
    clearYoutubePlayQuitOnDisconnectArmTimer,
    getQuitOnDisconnectArmed: (): boolean => quitOnDisconnectArmed,
    runYoutubePlaybackFlow,
  };
}
