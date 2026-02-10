import {
  OverlayShortcutFallbackHandlers,
} from "./overlay-shortcut-fallback-runner";
import { OverlayShortcutHandlers } from "./overlay-shortcut-service";

export interface OverlayShortcutRuntimeDeps {
  showMpvOsd: (text: string) => void;
  openRuntimeOptions: () => void;
  openJimaku: () => void;
  markAudioCard: () => Promise<void>;
  copySubtitleMultiple: (timeoutMs: number) => void;
  copySubtitle: () => void;
  toggleSecondarySub: () => void;
  updateLastCardFromClipboard: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsync: () => Promise<void>;
  mineSentence: () => Promise<void>;
  mineSentenceMultiple: (timeoutMs: number) => void;
}

function wrapAsync(
  task: () => Promise<void>,
  deps: OverlayShortcutRuntimeDeps,
  logLabel: string,
  osdLabel: string,
): () => void {
  return () => {
    task().catch((err) => {
      console.error(`${logLabel} failed:`, err);
      deps.showMpvOsd(`${osdLabel}: ${(err as Error).message}`);
    });
  };
}

export function createOverlayShortcutRuntimeHandlers(
  deps: OverlayShortcutRuntimeDeps,
): {
  overlayHandlers: OverlayShortcutHandlers;
  fallbackHandlers: OverlayShortcutFallbackHandlers;
} {
  const overlayHandlers: OverlayShortcutHandlers = {
    copySubtitle: () => {
      deps.copySubtitle();
    },
    copySubtitleMultiple: (timeoutMs) => {
      deps.copySubtitleMultiple(timeoutMs);
    },
    updateLastCardFromClipboard: wrapAsync(
      () => deps.updateLastCardFromClipboard(),
      deps,
      "updateLastCardFromClipboard",
      "Update failed",
    ),
    triggerFieldGrouping: wrapAsync(
      () => deps.triggerFieldGrouping(),
      deps,
      "triggerFieldGrouping",
      "Field grouping failed",
    ),
    triggerSubsync: wrapAsync(
      () => deps.triggerSubsync(),
      deps,
      "triggerSubsyncFromConfig",
      "Subsync failed",
    ),
    mineSentence: wrapAsync(
      () => deps.mineSentence(),
      deps,
      "mineSentenceCard",
      "Mine sentence failed",
    ),
    mineSentenceMultiple: (timeoutMs) => {
      deps.mineSentenceMultiple(timeoutMs);
    },
    toggleSecondarySub: () => deps.toggleSecondarySub(),
    markAudioCard: wrapAsync(
      () => deps.markAudioCard(),
      deps,
      "markLastCardAsAudioCard",
      "Audio card failed",
    ),
    openRuntimeOptions: () => {
      deps.openRuntimeOptions();
    },
    openJimaku: () => {
      deps.openJimaku();
    },
  };

  const fallbackHandlers: OverlayShortcutFallbackHandlers = {
    openRuntimeOptions: overlayHandlers.openRuntimeOptions,
    openJimaku: overlayHandlers.openJimaku,
    markAudioCard: overlayHandlers.markAudioCard,
    copySubtitleMultiple: overlayHandlers.copySubtitleMultiple,
    copySubtitle: overlayHandlers.copySubtitle,
    toggleSecondarySub: overlayHandlers.toggleSecondarySub,
    updateLastCardFromClipboard: overlayHandlers.updateLastCardFromClipboard,
    triggerFieldGrouping: overlayHandlers.triggerFieldGrouping,
    triggerSubsync: overlayHandlers.triggerSubsync,
    mineSentence: overlayHandlers.mineSentence,
    mineSentenceMultiple: overlayHandlers.mineSentenceMultiple,
  };

  return { overlayHandlers, fallbackHandlers };
}
