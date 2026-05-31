import type { SubtitleData } from '../../types';

export interface SubtitleProcessingControllerDeps {
  tokenizeSubtitle: (text: string) => Promise<SubtitleData | null>;
  emitSubtitle: (payload: SubtitleData) => void;
  logDebug?: (message: string) => void;
  now?: () => number;
}

export interface SubtitleProcessingController {
  onSubtitleChange: (text: string) => void;
  refreshCurrentSubtitle: (textOverride?: string) => void;
  invalidateTokenizationCache: () => void;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  consumeCachedSubtitle: (text: string) => SubtitleData | null;
  hasCachedSubtitle: (text: string) => boolean;
  isCacheFull: () => boolean;
}

export function normalizeSubtitleCacheKey(text: string): string {
  return text.replace(/\r\n/g, '\n').replace(/\\N/g, '\n').replace(/\\n/g, '\n').trim();
}

export function createSubtitleProcessingController(
  deps: SubtitleProcessingControllerDeps,
): SubtitleProcessingController {
  const SUBTITLE_TOKENIZATION_CACHE_LIMIT = 256;
  let latestText = '';
  let lastEmittedText = '';
  let cacheGeneration = 0;
  let lastEmittedGeneration = 0;
  let processing = false;
  let staleDropCount = 0;
  const tokenizationCache = new Map<string, SubtitleData>();
  const now = deps.now ?? (() => Date.now());

  const getCachedTokenization = (text: string): SubtitleData | null => {
    const cacheKey = normalizeSubtitleCacheKey(text);
    const cached = tokenizationCache.get(cacheKey);
    if (!cached) {
      return null;
    }

    tokenizationCache.delete(cacheKey);
    tokenizationCache.set(cacheKey, cached);
    return cached;
  };

  const setCachedTokenization = (text: string, payload: SubtitleData): void => {
    tokenizationCache.set(normalizeSubtitleCacheKey(text), payload);
    while (tokenizationCache.size > SUBTITLE_TOKENIZATION_CACHE_LIMIT) {
      const firstKey = tokenizationCache.keys().next().value;
      if (firstKey !== undefined) {
        tokenizationCache.delete(firstKey);
      }
    }
  };

  const processLatest = (): void => {
    if (processing) {
      return;
    }

    processing = true;

    void (async () => {
      while (true) {
        const text = latestText;
        const generation = cacheGeneration;
        const startedAtMs = now();

        if (!text.trim()) {
          deps.emitSubtitle({ text, tokens: null });
          lastEmittedText = text;
          lastEmittedGeneration = generation;
          break;
        }

        let output: SubtitleData = { text, tokens: null };
        try {
          const cachedTokenized = getCachedTokenization(text);
          if (cachedTokenized) {
            output = cachedTokenized;
          } else {
            const tokenized = await deps.tokenizeSubtitle(text);
            if (tokenized) {
              output = tokenized;
            }
            setCachedTokenization(text, output);
          }
        } catch (error) {
          deps.logDebug?.(`Subtitle tokenization failed: ${(error as Error).message}`);
        }

        if (latestText !== text) {
          staleDropCount += 1;
          deps.logDebug?.(
            `Dropped stale subtitle tokenization result; dropped=${staleDropCount}, elapsed=${now() - startedAtMs}ms`,
          );
          continue;
        }

        if (generation !== cacheGeneration) {
          deps.logDebug?.(
            `Dropped stale subtitle tokenization result after cache invalidation; elapsed=${now() - startedAtMs}ms`,
          );
          continue;
        }

        deps.emitSubtitle(output);
        lastEmittedText = text;
        lastEmittedGeneration = generation;
        deps.logDebug?.(
          `Subtitle tokenization delivered; elapsed=${now() - startedAtMs}ms, staleDrops=${staleDropCount}`,
        );
        break;
      }
    })()
      .catch((error) => {
        deps.logDebug?.(`Subtitle processing loop failed: ${(error as Error).message}`);
      })
      .finally(() => {
        processing = false;
        if (
          latestText !== lastEmittedText ||
          (latestText.trim() && cacheGeneration !== lastEmittedGeneration)
        ) {
          processLatest();
        }
      });
  };

  return {
    onSubtitleChange: (text: string) => {
      if (text === latestText) {
        return;
      }
      latestText = text;
      processLatest();
    },
    refreshCurrentSubtitle: (textOverride?: string) => {
      if (typeof textOverride === 'string') {
        latestText = textOverride;
      }
      if (!latestText.trim()) {
        return;
      }
      if (
        processing ||
        (latestText === lastEmittedText && cacheGeneration === lastEmittedGeneration)
      ) {
        return;
      }
      processLatest();
    },
    invalidateTokenizationCache: () => {
      tokenizationCache.clear();
      cacheGeneration += 1;
    },
    preCacheTokenization: (text: string, data: SubtitleData) => {
      setCachedTokenization(text, data);
    },
    consumeCachedSubtitle: (text: string) => {
      const cached = getCachedTokenization(text);
      if (!cached) {
        return null;
      }

      latestText = text;
      lastEmittedText = text;
      lastEmittedGeneration = cacheGeneration;
      return cached;
    },
    hasCachedSubtitle: (text: string) => {
      return tokenizationCache.has(normalizeSubtitleCacheKey(text));
    },
    isCacheFull: () => {
      return tokenizationCache.size >= SUBTITLE_TOKENIZATION_CACHE_LIMIT;
    },
  };
}
