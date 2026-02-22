import type { LauncherYoutubeSubgenConfig } from '../types.js';

function asStringArray(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined;
  return value.filter((entry): entry is string => typeof entry === 'string');
}

export function parseLauncherYoutubeSubgenConfig(
  root: Record<string, unknown>,
): LauncherYoutubeSubgenConfig {
  const youtubeSubgenRaw = root.youtubeSubgen;
  const youtubeSubgen =
    youtubeSubgenRaw && typeof youtubeSubgenRaw === 'object'
      ? (youtubeSubgenRaw as Record<string, unknown>)
      : null;
  const secondarySubRaw = root.secondarySub;
  const secondarySub =
    secondarySubRaw && typeof secondarySubRaw === 'object'
      ? (secondarySubRaw as Record<string, unknown>)
      : null;
  const jimakuRaw = root.jimaku;
  const jimaku =
    jimakuRaw && typeof jimakuRaw === 'object' ? (jimakuRaw as Record<string, unknown>) : null;

  const mode = youtubeSubgen?.mode;
  const jimakuLanguagePreference = jimaku?.languagePreference;
  const jimakuMaxEntryResults = jimaku?.maxEntryResults;

  return {
    mode: mode === 'automatic' || mode === 'preprocess' || mode === 'off' ? mode : undefined,
    whisperBin:
      typeof youtubeSubgen?.whisperBin === 'string' ? youtubeSubgen.whisperBin : undefined,
    whisperModel:
      typeof youtubeSubgen?.whisperModel === 'string' ? youtubeSubgen.whisperModel : undefined,
    primarySubLanguages: asStringArray(youtubeSubgen?.primarySubLanguages),
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
