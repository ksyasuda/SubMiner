import type { SubtitleData } from '../../types';

export interface SubtitleProcessingControllerDeps {
  tokenizeSubtitle: (text: string) => Promise<SubtitleData | null>;
  emitSubtitle: (payload: SubtitleData) => void;
  logDebug?: (message: string) => void;
  now?: () => number;
  cacheLimit?: number;
}

/**
 * Pure memory bound on the LRU, not a coverage limit: prefetching runs to the end of a
 * file regardless of cache pressure. Sized to hold a feature-length title (a 24-minute
 * episode runs 300-400 lines, a 2-hour film ~2000) plus room for lines that repeat across
 * episodes of a series, so openings and endings stay warm between titles.
 */
export const DEFAULT_SUBTITLE_TOKENIZATION_CACHE_LIMIT = 2500;

export interface SubtitleProcessingController {
  /**
   * Returns whether the text was new and processing was scheduled. A false
   * return means nothing will be emitted for this event, which callers that
   * gate work on the emit (such as pausing subtitle prefetching) need to know.
   */
  onSubtitleChange: (text: string) => boolean;
  /** Same contract as onSubtitleChange: whether an emit is expected. */
  refreshCurrentSubtitle: (textOverride?: string) => boolean;
  invalidateTokenizationCache: () => void;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  consumeCachedSubtitle: (text: string) => SubtitleData | null;
  hasCachedSubtitle: (text: string) => boolean;
}

export function normalizeSubtitleCacheKey(text: string): string {
  return text.replace(/\r\n/g, '\n').replace(/\\N/g, '\n').replace(/\\n/g, '\n').trim();
}

export function createSubtitleProcessingController(
  deps: SubtitleProcessingControllerDeps,
): SubtitleProcessingController {
  const SUBTITLE_TOKENIZATION_CACHE_LIMIT =
    deps.cacheLimit && deps.cacheLimit > 0
      ? deps.cacheLimit
      : DEFAULT_SUBTITLE_TOKENIZATION_CACHE_LIMIT;
  let latestText = '';
  let lastEmittedText = '';
  // Tracks the latest provisional plain emit across rapid changes and loop retries
  // so the same line is never shown plain twice.
  let lastPlainEmittedText: string | null = null;
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
          if (lastPlainEmittedText !== text) {
            deps.emitSubtitle({ text, tokens: null });
          }
          lastEmittedText = text;
          lastEmittedGeneration = generation;
          lastPlainEmittedText = null;
          break;
        }

        let output: SubtitleData = { text, tokens: null };
        try {
          const cachedTokenized = getCachedTokenization(text);
          if (cachedTokenized) {
            output = cachedTokenized;
          } else {
            // Cache miss: show the plain line on time; the tokenized payload
            // upgrades it once ready. Skipped on refreshes of an already
            // emitted line so downstream consumers never see a downgrade.
            if (text !== lastEmittedText && text !== lastPlainEmittedText) {
              deps.emitSubtitle({ text, tokens: null });
              lastPlainEmittedText = text;
            }
            const tokenized = await deps.tokenizeSubtitle(text);
            // A null result is a transient tokenizer failure, not a verdict on
            // the line: caching the plain fallback would pin it untokenized for
            // every later occurrence.
            if (tokenized) {
              output = tokenized;
              // A result computed before an invalidation must not repopulate the
              // fresh cache, or the retry below would serve the stale entry.
              if (generation === cacheGeneration) {
                setCachedTokenization(text, tokenized);
              }
            }
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

        // An untokenized result adds nothing when this line was already shown,
        // either provisionally or as an earlier full emit (failed refresh) —
        // emitting it would duplicate or downgrade what is on screen.
        const plainAlreadyShown = lastPlainEmittedText === text || lastEmittedText === text;
        if (!(output.tokens === null && output.text === text && plainAlreadyShown)) {
          deps.emitSubtitle(output);
        }
        lastEmittedText = text;
        lastEmittedGeneration = generation;
        lastPlainEmittedText = null;
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
        // A run already in flight for this text will still emit for it.
        return processing;
      }
      latestText = text;
      if (
        processing &&
        text !== lastPlainEmittedText &&
        !tokenizationCache.has(normalizeSubtitleCacheKey(text))
      ) {
        deps.emitSubtitle({ text, tokens: null });
        lastPlainEmittedText = text;
      }
      processLatest();
      return true;
    },
    refreshCurrentSubtitle: (textOverride?: string) => {
      if (typeof textOverride === 'string') {
        latestText = textOverride;
      }
      if (!latestText.trim()) {
        // A run in flight will pick this up and emit the empty subtitle, so
        // the caller is still waiting on an emit.
        return processing;
      }
      if (processing) {
        return true;
      }
      if (latestText === lastEmittedText && cacheGeneration === lastEmittedGeneration) {
        return false;
      }
      processLatest();
      return true;
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
      lastPlainEmittedText = null;
      return cached;
    },
    hasCachedSubtitle: (text: string) => {
      return tokenizationCache.has(normalizeSubtitleCacheKey(text));
    },
  };
}
