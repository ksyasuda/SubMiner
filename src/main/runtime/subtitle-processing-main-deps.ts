import type { SubtitleProcessingControllerDeps } from '../../core/services/subtitle-processing-controller';

export function createBuildSubtitleProcessingControllerMainDepsHandler(
  deps: SubtitleProcessingControllerDeps,
) {
  return (): SubtitleProcessingControllerDeps => ({
    tokenizeSubtitle: (text: string) => deps.tokenizeSubtitle(text),
    emitSubtitle: (payload) => deps.emitSubtitle(payload),
    logDebug: deps.logDebug,
    now: deps.now,
  });
}
