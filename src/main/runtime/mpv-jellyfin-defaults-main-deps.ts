import type {
  createApplyJellyfinMpvDefaultsHandler,
  createGetDefaultSocketPathHandler,
} from './mpv-jellyfin-defaults';

type ApplyJellyfinMpvDefaultsMainDeps = Parameters<typeof createApplyJellyfinMpvDefaultsHandler>[0];
type GetDefaultSocketPathMainDeps = Parameters<typeof createGetDefaultSocketPathHandler>[0];

export function createBuildApplyJellyfinMpvDefaultsMainDepsHandler(
  deps: ApplyJellyfinMpvDefaultsMainDeps,
) {
  return (): ApplyJellyfinMpvDefaultsMainDeps => ({
    sendMpvCommandRuntime: (client, command) => deps.sendMpvCommandRuntime(client, command),
    jellyfinLangPref: deps.jellyfinLangPref,
  });
}

export function createBuildGetDefaultSocketPathMainDepsHandler(
  deps: GetDefaultSocketPathMainDeps,
) {
  return (): GetDefaultSocketPathMainDeps => ({
    platform: deps.platform,
  });
}
