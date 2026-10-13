import type { LauncherYoutubeSubgenConfig } from '../types.js';

function asStringArray(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined;
  return value.filter((entry): entry is string => typeof entry === 'string');
}

export function parseLauncherYoutubeSubgenConfig(
  root: Record<string, unknown>,
): LauncherYoutubeSubgenConfig {
  const youtubeRaw = root.youtube;
  const youtube =
    youtubeRaw && typeof youtubeRaw === 'object' ? (youtubeRaw as Record<string, unknown>) : null;
  const secondarySubRaw = root.secondarySub;
  const secondarySub =
    secondarySubRaw && typeof secondarySubRaw === 'object'
      ? (secondarySubRaw as Record<string, unknown>)
      : null;

  return {
    primarySubLanguages: asStringArray(youtube?.primarySubLanguages),
    secondarySubLanguages: asStringArray(secondarySub?.secondarySubLanguages),
  };
}
