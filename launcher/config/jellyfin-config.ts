import type { LauncherJellyfinConfig } from '../types.js';

export function parseLauncherJellyfinConfig(root: Record<string, unknown>): LauncherJellyfinConfig {
  const jellyfinRaw = root.jellyfin;
  if (!jellyfinRaw || typeof jellyfinRaw !== 'object') return {};
  const jellyfin = jellyfinRaw as Record<string, unknown>;
  return {
    enabled: typeof jellyfin.enabled === 'boolean' ? jellyfin.enabled : undefined,
    serverUrl: typeof jellyfin.serverUrl === 'string' ? jellyfin.serverUrl : undefined,
    username: typeof jellyfin.username === 'string' ? jellyfin.username : undefined,
    defaultLibraryId:
      typeof jellyfin.defaultLibraryId === 'string' ? jellyfin.defaultLibraryId : undefined,
    pullPictures: typeof jellyfin.pullPictures === 'boolean' ? jellyfin.pullPictures : undefined,
    iconCacheDir: typeof jellyfin.iconCacheDir === 'string' ? jellyfin.iconCacheDir : undefined,
  };
}
