import type {
  SubtitlePrefetchService,
  SubtitlePrefetchServiceDeps,
} from '../../core/services/subtitle-prefetch';
import type { SubtitleData } from '../../types';
import type { SubtitleCue } from '../../types';
import { buildOverlapPrefetchCues } from './subtitle-prefetch-overlaps';

export interface SubtitlePrefetchInitControllerDeps {
  getCurrentService: () => SubtitlePrefetchService | null;
  setCurrentService: (service: SubtitlePrefetchService | null) => void;
  loadSubtitleSourceText: (source: string) => Promise<string>;
  parseSubtitleCues: (content: string, filename: string) => SubtitleCue[];
  createSubtitlePrefetchService: (deps: SubtitlePrefetchServiceDeps) => SubtitlePrefetchService;
  tokenizeSubtitle: (text: string) => Promise<SubtitleData | null>;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  hasCachedTokenization?: (text: string) => boolean;
  getCacheGeneration?: () => number;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
  onParsedSubtitleCuesChanged?: (cues: SubtitleCue[] | null, sourceKey: string | null) => void;
}

export interface SubtitlePrefetchInitController {
  cancelPendingInit: () => void;
  initSubtitlePrefetch: (
    sourcePath: string,
    currentTimePos: number,
    sourceKey?: string,
  ) => Promise<void>;
}

export function createSubtitlePrefetchInitController(
  deps: SubtitlePrefetchInitControllerDeps,
): SubtitlePrefetchInitController {
  let initRevision = 0;
  // What the running service was built from. Startup fires several inits for one source
  // (Jellyfin preload seed, sid change, each track-list update); restarting on each would
  // discard the in-flight tokenization and re-queue the priority window behind it.
  let activeSource: { key: string; content: string; cues: SubtitleCue[] } | null = null;

  const stopCurrentService = (): void => {
    deps.getCurrentService()?.stop();
    deps.setCurrentService(null);
    activeSource = null;
  };

  const cancelPendingInit = (): void => {
    initRevision += 1;
    stopCurrentService();
    deps.onParsedSubtitleCuesChanged?.(null, null);
  };

  const initSubtitlePrefetch = async (
    sourcePath: string,
    currentTimePos: number,
    sourceKey = sourcePath,
  ): Promise<void> => {
    const revision = ++initRevision;
    // The same source may still be current once loaded; keep its service running until then.
    if (activeSource?.key !== sourceKey) {
      stopCurrentService();
    }

    try {
      const content = await deps.loadSubtitleSourceText(sourcePath);
      if (revision !== initRevision) {
        return;
      }
      if (
        activeSource?.key === sourceKey &&
        activeSource.content === content &&
        deps.getCurrentService()
      ) {
        // A track change clears the published cues synchronously, so republish them.
        deps.onParsedSubtitleCuesChanged?.(activeSource.cues, sourceKey);
        return;
      }
      stopCurrentService();

      const cues = deps.parseSubtitleCues(content, sourcePath);
      if (revision !== initRevision || cues.length === 0) {
        if (revision === initRevision) {
          deps.logWarn(
            `[subtitle-prefetch] parsed 0 cues from ${sourcePath}; prefetch disabled for this source`,
          );
          deps.onParsedSubtitleCuesChanged?.(null, null);
        }
        return;
      }

      // Overlap cues only feed the cache; the published cue list stays as authored.
      const prefetchCues = [...cues, ...buildOverlapPrefetchCues(cues)].sort(
        (a, b) => a.startTime - b.startTime,
      );
      const nextService = deps.createSubtitlePrefetchService({
        cues: prefetchCues,
        tokenizeSubtitle: (text) => deps.tokenizeSubtitle(text),
        preCacheTokenization: (text, data) => deps.preCacheTokenization(text, data),
        hasCachedTokenization: (text) => deps.hasCachedTokenization?.(text) ?? false,
        getCacheGeneration: deps.getCacheGeneration,
      });

      if (revision !== initRevision) {
        return;
      }

      deps.setCurrentService(nextService);
      activeSource = { key: sourceKey, content, cues };
      deps.onParsedSubtitleCuesChanged?.(cues, sourceKey);
      nextService.start(currentTimePos);
      deps.logInfo(
        `[subtitle-prefetch] started prefetching ${cues.length} cues from ${sourcePath}`,
      );
    } catch (error) {
      if (revision === initRevision) {
        stopCurrentService();
        deps.onParsedSubtitleCuesChanged?.(null, null);
        deps.logWarn(`[subtitle-prefetch] failed to initialize: ${(error as Error).message}`);
      }
    }
  };

  return {
    cancelPendingInit,
    initSubtitlePrefetch,
  };
}
