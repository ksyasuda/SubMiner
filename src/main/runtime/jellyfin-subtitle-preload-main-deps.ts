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
    logDebug: (message: string, error: unknown) => deps.logDebug(message, error),
  });
}
