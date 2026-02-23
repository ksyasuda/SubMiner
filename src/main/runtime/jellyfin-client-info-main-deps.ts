import type {
  createGetJellyfinClientInfoHandler,
  createGetResolvedJellyfinConfigHandler,
} from './jellyfin-client-info';

type GetResolvedJellyfinConfigMainDeps = Parameters<typeof createGetResolvedJellyfinConfigHandler>[0];
type GetJellyfinClientInfoMainDeps = Parameters<typeof createGetJellyfinClientInfoHandler>[0];

export function createBuildGetResolvedJellyfinConfigMainDepsHandler(
  deps: GetResolvedJellyfinConfigMainDeps,
) {
  return (): GetResolvedJellyfinConfigMainDeps => ({
    getResolvedConfig: () => deps.getResolvedConfig(),
    loadStoredSession: () => deps.loadStoredSession(),
    getEnv: (name: string) => deps.getEnv(name),
  });
}

export function createBuildGetJellyfinClientInfoMainDepsHandler(
  deps: GetJellyfinClientInfoMainDeps,
) {
  return (): GetJellyfinClientInfoMainDeps => ({
    getResolvedJellyfinConfig: () => deps.getResolvedJellyfinConfig(),
    getDefaultJellyfinConfig: () => deps.getDefaultJellyfinConfig(),
  });
}
