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
    getSavedSubtitleDelay: deps.getSavedSubtitleDelay
      ? (itemId, streamIndex) => deps.getSavedSubtitleDelay!(itemId, streamIndex)
      : undefined,
    setActiveSubtitleDelayKey: deps.setActiveSubtitleDelayKey
      ? (key) => deps.setActiveSubtitleDelayKey!(key)
      : undefined,
    loadSubtitleSourceText: deps.loadSubtitleSourceText
      ? (source) => deps.loadSubtitleSourceText!(source)
      : undefined,
    saveSubtitleDelay: deps.saveSubtitleDelay
      ? (itemId, streamIndex, delaySeconds) =>
          deps.saveSubtitleDelay!(itemId, streamIndex, delaySeconds)
      : undefined,
    initSubtitlePrefetch: deps.initSubtitlePrefetch
      ? (sourcePath) => deps.initSubtitlePrefetch!(sourcePath)
      : undefined,
    logDebug: (message: string, error: unknown) => deps.logDebug(message, error),
  });
}
