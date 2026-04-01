import {
  createBuildGetDefaultSocketPathMainDepsHandler,
  createGetDefaultSocketPathHandler,
} from './runtime/domains/jellyfin';

export function createDefaultSocketPathResolver(platform: NodeJS.Platform) {
  return createGetDefaultSocketPathHandler(
    createBuildGetDefaultSocketPathMainDepsHandler({
      platform,
    })(),
  );
}
