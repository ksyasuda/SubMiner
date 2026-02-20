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
    loadStoredToken: () => deps.loadStoredToken(),
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
