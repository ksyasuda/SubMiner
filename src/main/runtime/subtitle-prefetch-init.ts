import type {
  SubtitlePrefetchService,
  SubtitlePrefetchServiceDeps,
} from '../../core/services/subtitle-prefetch';
import type { SubtitleData } from '../../types';
import type { SubtitleCue } from '../../types';

export interface SubtitlePrefetchInitControllerDeps {
  getCurrentService: () => SubtitlePrefetchService | null;
  setCurrentService: (service: SubtitlePrefetchService | null) => void;
  loadSubtitleSourceText: (source: string) => Promise<string>;
  parseSubtitleCues: (content: string, filename: string) => SubtitleCue[];
  createSubtitlePrefetchService: (deps: SubtitlePrefetchServiceDeps) => SubtitlePrefetchService;
  tokenizeSubtitle: (text: string) => Promise<SubtitleData | null>;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  hasCachedTokenization?: (text: string) => boolean;
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

  const cancelPendingInit = (): void => {
    initRevision += 1;
    deps.getCurrentService()?.stop();
    deps.setCurrentService(null);
    deps.onParsedSubtitleCuesChanged?.(null, null);
  };

  const initSubtitlePrefetch = async (
    sourcePath: string,
    currentTimePos: number,
    sourceKey = sourcePath,
  ): Promise<void> => {
    const revision = ++initRevision;
    deps.getCurrentService()?.stop();
    deps.setCurrentService(null);

    try {
      const content = await deps.loadSubtitleSourceText(sourcePath);
      if (revision !== initRevision) {
        return;
      }

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

      const nextService = deps.createSubtitlePrefetchService({
        cues,
        tokenizeSubtitle: (text) => deps.tokenizeSubtitle(text),
        preCacheTokenization: (text, data) => deps.preCacheTokenization(text, data),
        hasCachedTokenization: (text) => deps.hasCachedTokenization?.(text) ?? false,
      });

      if (revision !== initRevision) {
        return;
      }

      deps.setCurrentService(nextService);
      deps.onParsedSubtitleCuesChanged?.(cues, sourceKey);
      nextService.start(currentTimePos);
      deps.logInfo(
        `[subtitle-prefetch] started prefetching ${cues.length} cues from ${sourcePath}`,
      );
    } catch (error) {
      if (revision === initRevision) {
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
