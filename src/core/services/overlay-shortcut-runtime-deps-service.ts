import {
  OverlayShortcutLifecycleDeps,
} from "./overlay-shortcut-lifecycle-service";
import {
  OverlayShortcutRuntimeDeps,
} from "./overlay-shortcut-runtime-service";

export interface OverlayShortcutRuntimeDepsOptions {
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

export interface OverlayShortcutLifecycleDepsOptions {
  getConfiguredShortcuts: OverlayShortcutLifecycleDeps["getConfiguredShortcuts"];
  getOverlayHandlers: OverlayShortcutLifecycleDeps["getOverlayHandlers"];
  cancelPendingMultiCopy: () => void;
  cancelPendingMineSentenceMultiple: () => void;
}

export function createOverlayShortcutRuntimeDepsService(
  options: OverlayShortcutRuntimeDepsOptions,
): OverlayShortcutRuntimeDeps {
  return {
    showMpvOsd: options.showMpvOsd,
    openRuntimeOptions: options.openRuntimeOptions,
    openJimaku: options.openJimaku,
    markAudioCard: options.markAudioCard,
    copySubtitleMultiple: options.copySubtitleMultiple,
    copySubtitle: options.copySubtitle,
    toggleSecondarySub: options.toggleSecondarySub,
    updateLastCardFromClipboard: options.updateLastCardFromClipboard,
    triggerFieldGrouping: options.triggerFieldGrouping,
    triggerSubsync: options.triggerSubsync,
    mineSentence: options.mineSentence,
    mineSentenceMultiple: options.mineSentenceMultiple,
  };
}

export function createOverlayShortcutLifecycleDepsRuntimeService(
  options: OverlayShortcutLifecycleDepsOptions,
): OverlayShortcutLifecycleDeps {
  return {
    getConfiguredShortcuts: options.getConfiguredShortcuts,
    getOverlayHandlers: options.getOverlayHandlers,
    cancelPendingMultiCopy: options.cancelPendingMultiCopy,
    cancelPendingMineSentenceMultiple:
      options.cancelPendingMineSentenceMultiple,
  };
}
