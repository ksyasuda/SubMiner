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
  const jimakuRaw = root.jimaku;
  const jimaku =
    jimakuRaw && typeof jimakuRaw === 'object' ? (jimakuRaw as Record<string, unknown>) : null;

  const jimakuLanguagePreference = jimaku?.languagePreference;
  const jimakuMaxEntryResults = jimaku?.maxEntryResults;

  return {
    primarySubLanguages: asStringArray(youtube?.primarySubLanguages),
    secondarySubLanguages: asStringArray(secondarySub?.secondarySubLanguages),
    jimakuApiKey: typeof jimaku?.apiKey === 'string' ? jimaku.apiKey : undefined,
    jimakuApiKeyCommand:
      typeof jimaku?.apiKeyCommand === 'string' ? jimaku.apiKeyCommand : undefined,
    jimakuApiBaseUrl: typeof jimaku?.apiBaseUrl === 'string' ? jimaku.apiBaseUrl : undefined,
    jimakuLanguagePreference:
      jimakuLanguagePreference === 'ja' ||
      jimakuLanguagePreference === 'en' ||
      jimakuLanguagePreference === 'none'
        ? jimakuLanguagePreference
        : undefined,
    jimakuMaxEntryResults:
      typeof jimakuMaxEntryResults === 'number' &&
      Number.isFinite(jimakuMaxEntryResults) &&
      jimakuMaxEntryResults > 0
        ? Math.floor(jimakuMaxEntryResults)
        : undefined,
  };
}
