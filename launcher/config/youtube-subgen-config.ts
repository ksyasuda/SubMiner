import type { LauncherYoutubeSubgenConfig } from '../types.js';
import { mergeAiConfig } from '../../src/ai/config.js';

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
  const aiRaw = root.ai;
  const ai = aiRaw && typeof aiRaw === 'object' ? (aiRaw as Record<string, unknown>) : null;
  const youtubeAiRaw = youtubeSubgen?.ai;
  const youtubeAi =
    youtubeAiRaw && typeof youtubeAiRaw === 'object'
      ? (youtubeAiRaw as Record<string, unknown>)
      : null;

  const jimakuLanguagePreference = jimaku?.languagePreference;
  const jimakuMaxEntryResults = jimaku?.maxEntryResults;

  return {
    whisperBin:
      typeof youtubeSubgen?.whisperBin === 'string' ? youtubeSubgen.whisperBin : undefined,
    whisperModel:
      typeof youtubeSubgen?.whisperModel === 'string' ? youtubeSubgen.whisperModel : undefined,
    whisperVadModel:
      typeof youtubeSubgen?.whisperVadModel === 'string'
        ? youtubeSubgen.whisperVadModel
        : undefined,
    whisperThreads:
      typeof youtubeSubgen?.whisperThreads === 'number' &&
      Number.isFinite(youtubeSubgen.whisperThreads) &&
      youtubeSubgen.whisperThreads > 0
        ? Math.floor(youtubeSubgen.whisperThreads)
        : undefined,
    fixWithAi: typeof youtubeSubgen?.fixWithAi === 'boolean' ? youtubeSubgen.fixWithAi : undefined,
    ai: mergeAiConfig(
      ai
        ? {
            enabled: typeof ai.enabled === 'boolean' ? ai.enabled : undefined,
            apiKey: typeof ai.apiKey === 'string' ? ai.apiKey : undefined,
            apiKeyCommand: typeof ai.apiKeyCommand === 'string' ? ai.apiKeyCommand : undefined,
            baseUrl: typeof ai.baseUrl === 'string' ? ai.baseUrl : undefined,
            model: typeof ai.model === 'string' ? ai.model : undefined,
            systemPrompt: typeof ai.systemPrompt === 'string' ? ai.systemPrompt : undefined,
            requestTimeoutMs:
              typeof ai.requestTimeoutMs === 'number' &&
              Number.isFinite(ai.requestTimeoutMs) &&
              ai.requestTimeoutMs > 0
                ? Math.floor(ai.requestTimeoutMs)
                : undefined,
          }
        : undefined,
      youtubeAi
        ? {
            model: typeof youtubeAi.model === 'string' ? youtubeAi.model : undefined,
            systemPrompt:
              typeof youtubeAi.systemPrompt === 'string' ? youtubeAi.systemPrompt : undefined,
          }
        : undefined,
    ),
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
