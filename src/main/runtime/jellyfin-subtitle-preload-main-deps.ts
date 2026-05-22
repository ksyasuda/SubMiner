import type { createPreloadJellyfinExternalSubtitlesHandler } from './jellyfin-subtitle-preload';

type PreloadJellyfinExternalSubtitlesMainDeps = Parameters<
  typeof createPreloadJellyfinExternalSubtitlesHandler
>[0];

export function createBuildPreloadJellyfinExternalSubtitlesMainDepsHandler(
  deps: PreloadJellyfinExternalSubtitlesMainDeps,
) {
  return (): PreloadJellyfinExternalSubtitlesMainDeps => ({
    listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
      deps.listJellyfinSubtitleTracks(session, clientInfo, itemId),
    getMpvClient: () => deps.getMpvClient(),
    sendMpvCommand: (command) => deps.sendMpvCommand(command),
    wait: (ms: number) => deps.wait(ms),
    cacheSubtitleTrack: (track) => deps.cacheSubtitleTrack(track),
    cleanupCachedSubtitles: (dirs) => deps.cleanupCachedSubtitles(dirs),
    logDebug: (message: string, error: unknown) => deps.logDebug(message, error),
  });
}
