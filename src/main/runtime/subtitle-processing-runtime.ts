import {
  createSubtitleProcessingController,
  type SubtitleProcessingControllerDeps,
} from '../../core/services/subtitle-processing-controller';
import type { SubtitleData } from '../../types';

/**
 * Creates the main-process subtitle processing controller and ties it to the
 * subtitle prefetch pause.
 *
 * Controller emits never release the prefetch pause: the first emit for an
 * uncached line is the provisional plain payload, sent before tokenization
 * starts, so resuming on it would put prefetch back in contention with the
 * on-screen line for the single parser window. The pause is released once the
 * controller has no work left, which covers the runs that end without an emit
 * (suppressed duplicate, failed tokenization) as well as the ones that deliver
 * a payload.
 */
export function createPrefetchAwareSubtitleProcessingController(deps: {
  tokenizeSubtitle: SubtitleProcessingControllerDeps['tokenizeSubtitle'];
  emitSubtitlePayload: (payload: SubtitleData, options: { resumePrefetch: boolean }) => void;
  resumePrefetch: () => void;
  logDebug: (message: string) => void;
  now: () => number;
}) {
  return createSubtitleProcessingController({
    tokenizeSubtitle: deps.tokenizeSubtitle,
    emitSubtitle: (payload) => deps.emitSubtitlePayload(payload, { resumePrefetch: false }),
    onProcessingSettled: () => deps.resumePrefetch(),
    logDebug: deps.logDebug,
    now: deps.now,
  });
}
