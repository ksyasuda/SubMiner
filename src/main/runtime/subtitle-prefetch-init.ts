import type { SubtitleCue } from '../../core/services/subtitle-cue-parser';
import type {
  SubtitlePrefetchService,
  SubtitlePrefetchServiceDeps,
} from '../../core/services/subtitle-prefetch';
import type { SubtitleData } from '../../types';

export interface SubtitlePrefetchInitControllerDeps {
  getCurrentService: () => SubtitlePrefetchService | null;
  setCurrentService: (service: SubtitlePrefetchService | null) => void;
  loadSubtitleSourceText: (source: string) => Promise<string>;
  parseSubtitleCues: (content: string, filename: string) => SubtitleCue[];
  createSubtitlePrefetchService: (deps: SubtitlePrefetchServiceDeps) => SubtitlePrefetchService;
  tokenizeSubtitle: (text: string) => Promise<SubtitleData | null>;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  isCacheFull: () => boolean;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
}

export interface SubtitlePrefetchInitController {
  cancelPendingInit: () => void;
  initSubtitlePrefetch: (externalFilename: string, currentTimePos: number) => Promise<void>;
}

export function createSubtitlePrefetchInitController(
  deps: SubtitlePrefetchInitControllerDeps,
): SubtitlePrefetchInitController {
  let initRevision = 0;

  const cancelPendingInit = (): void => {
    initRevision += 1;
    deps.getCurrentService()?.stop();
    deps.setCurrentService(null);
  };

  const initSubtitlePrefetch = async (
    externalFilename: string,
    currentTimePos: number,
  ): Promise<void> => {
    const revision = ++initRevision;
    deps.getCurrentService()?.stop();
    deps.setCurrentService(null);

    try {
      const content = await deps.loadSubtitleSourceText(externalFilename);
      if (revision !== initRevision) {
        return;
      }

      const cues = deps.parseSubtitleCues(content, externalFilename);
      if (revision !== initRevision || cues.length === 0) {
        return;
      }

      const nextService = deps.createSubtitlePrefetchService({
        cues,
        tokenizeSubtitle: (text) => deps.tokenizeSubtitle(text),
        preCacheTokenization: (text, data) => deps.preCacheTokenization(text, data),
        isCacheFull: () => deps.isCacheFull(),
      });

      if (revision !== initRevision) {
        return;
      }

      deps.setCurrentService(nextService);
      nextService.start(currentTimePos);
      deps.logInfo(
        `[subtitle-prefetch] started prefetching ${cues.length} cues from ${externalFilename}`,
      );
    } catch (error) {
      if (revision === initRevision) {
        deps.logWarn(`[subtitle-prefetch] failed to initialize: ${(error as Error).message}`);
      }
    }
  };

  return {
    cancelPendingInit,
    initSubtitlePrefetch,
  };
}
