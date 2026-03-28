import type { SubtitlePrefetchInitController } from '../subtitle-prefetch-init';
import type { ComposerInputs, ComposerOutputs } from './contracts';

export type SubtitlePrefetchRuntimeComposerOptions = ComposerInputs<{
  subtitlePrefetchInitController: SubtitlePrefetchInitController;
  refreshSubtitleSidebarFromSource: (sourcePath: string) => Promise<void>;
  refreshSubtitlePrefetchFromActiveTrack: () => Promise<void>;
  scheduleSubtitlePrefetchRefresh: (delayMs?: number) => void;
  clearScheduledSubtitlePrefetchRefresh: () => void;
}>;

export type SubtitlePrefetchRuntimeComposerResult = ComposerOutputs<{
  cancelPendingInit: () => void;
  initSubtitlePrefetch: SubtitlePrefetchInitController['initSubtitlePrefetch'];
  refreshSubtitleSidebarFromSource: (sourcePath: string) => Promise<void>;
  refreshSubtitlePrefetchFromActiveTrack: () => Promise<void>;
  scheduleSubtitlePrefetchRefresh: (delayMs?: number) => void;
  clearScheduledSubtitlePrefetchRefresh: () => void;
}>;

export function composeSubtitlePrefetchRuntime(
  options: SubtitlePrefetchRuntimeComposerOptions,
): SubtitlePrefetchRuntimeComposerResult {
  return {
    cancelPendingInit: () => options.subtitlePrefetchInitController.cancelPendingInit(),
    initSubtitlePrefetch: options.subtitlePrefetchInitController.initSubtitlePrefetch,
    refreshSubtitleSidebarFromSource: options.refreshSubtitleSidebarFromSource,
    refreshSubtitlePrefetchFromActiveTrack: options.refreshSubtitlePrefetchFromActiveTrack,
    scheduleSubtitlePrefetchRefresh: options.scheduleSubtitlePrefetchRefresh,
    clearScheduledSubtitlePrefetchRefresh: options.clearScheduledSubtitlePrefetchRefresh,
  };
}
