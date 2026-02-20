import type { createPlayJellyfinItemInMpvHandler } from './jellyfin-playback-launch';

type PlayJellyfinItemInMpvMainDeps = Parameters<typeof createPlayJellyfinItemInMpvHandler>[0];

export function createBuildPlayJellyfinItemInMpvMainDepsHandler(deps: PlayJellyfinItemInMpvMainDeps) {
  return (): PlayJellyfinItemInMpvMainDeps => ({
    ensureMpvConnectedForPlayback: () => deps.ensureMpvConnectedForPlayback(),
    getMpvClient: () => deps.getMpvClient(),
    resolvePlaybackPlan: (params) => deps.resolvePlaybackPlan(params),
    applyJellyfinMpvDefaults: (mpvClient) => deps.applyJellyfinMpvDefaults(mpvClient),
    sendMpvCommand: (command: Array<string | number>) => deps.sendMpvCommand(command),
    armQuitOnDisconnect: () => deps.armQuitOnDisconnect(),
    schedule: (callback: () => void, delayMs: number) => deps.schedule(callback, delayMs),
    convertTicksToSeconds: (ticks: number) => deps.convertTicksToSeconds(ticks),
    preloadExternalSubtitles: (params) => deps.preloadExternalSubtitles(params),
    setActivePlayback: (state) => deps.setActivePlayback(state),
    setLastProgressAtMs: (value: number) => deps.setLastProgressAtMs(value),
    reportPlaying: (payload) => deps.reportPlaying(payload),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
  });
}
