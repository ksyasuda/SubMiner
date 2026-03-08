import type { ConfiguredShortcuts } from '../core/utils/shortcut-config';
import {
  createOverlayShortcutRuntimeHandlers,
  shortcutMatchesInputForLocalFallback,
} from '../core/services';
import {
  refreshOverlayShortcutsRuntime,
  registerOverlayShortcuts,
  shouldActivateOverlayShortcuts,
  syncOverlayShortcutsRuntime,
  unregisterOverlayShortcutsRuntime,
} from '../core/services/overlay-shortcut';
import { runOverlayShortcutLocalFallback } from '../core/services/overlay-shortcut-handler';

export interface OverlayShortcutRuntimeServiceInput {
  getConfiguredShortcuts: () => ConfiguredShortcuts;
  getShortcutsRegistered: () => boolean;
  setShortcutsRegistered: (registered: boolean) => void;
  isOverlayRuntimeInitialized: () => boolean;
  isMacOSPlatform: () => boolean;
  isTrackedMpvWindowFocused: () => boolean;
  showMpvOsd: (text: string) => void;
  openRuntimeOptionsPalette: () => void;
  openJimaku: () => void;
  markAudioCard: () => Promise<void>;
  copySubtitleMultiple: (timeoutMs: number) => void;
  copySubtitle: () => void;
  toggleSecondarySubMode: () => void;
  updateLastCardFromClipboard: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsyncFromConfig: () => Promise<void>;
  mineSentenceCard: () => Promise<void>;
  mineSentenceMultiple: (timeoutMs: number) => void;
  cancelPendingMultiCopy: () => void;
  cancelPendingMineSentenceMultiple: () => void;
}

export interface OverlayShortcutsRuntimeService {
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  registerOverlayShortcuts: () => void;
  unregisterOverlayShortcuts: () => void;
  syncOverlayShortcuts: () => void;
  refreshOverlayShortcuts: () => void;
}

export function createOverlayShortcutsRuntimeService(
  input: OverlayShortcutRuntimeServiceInput,
): OverlayShortcutsRuntimeService {
  const handlers = createOverlayShortcutRuntimeHandlers({
    showMpvOsd: (text: string) => input.showMpvOsd(text),
    openRuntimeOptions: () => {
      input.openRuntimeOptionsPalette();
    },
    openJimaku: () => {
      input.openJimaku();
    },
    markAudioCard: () => {
      return input.markAudioCard();
    },
    copySubtitleMultiple: (timeoutMs: number) => {
      input.copySubtitleMultiple(timeoutMs);
    },
    copySubtitle: () => {
      input.copySubtitle();
    },
    toggleSecondarySub: () => {
      input.toggleSecondarySubMode();
    },
    updateLastCardFromClipboard: () => {
      return input.updateLastCardFromClipboard();
    },
    triggerFieldGrouping: () => {
      return input.triggerFieldGrouping();
    },
    triggerSubsync: () => {
      return input.triggerSubsyncFromConfig();
    },
    mineSentence: () => {
      return input.mineSentenceCard();
    },
    mineSentenceMultiple: (timeoutMs: number) => {
      input.mineSentenceMultiple(timeoutMs);
    },
  });

  const getShortcutLifecycleDeps = () => {
    return {
      getConfiguredShortcuts: () => input.getConfiguredShortcuts(),
      getOverlayHandlers: () => handlers.overlayHandlers,
      cancelPendingMultiCopy: () => input.cancelPendingMultiCopy(),
      cancelPendingMineSentenceMultiple: () => input.cancelPendingMineSentenceMultiple(),
    };
  };

  const shouldOverlayShortcutsBeActive = () =>
    shouldActivateOverlayShortcuts({
      overlayRuntimeInitialized: input.isOverlayRuntimeInitialized(),
      isMacOSPlatform: input.isMacOSPlatform(),
      trackedMpvWindowFocused: input.isTrackedMpvWindowFocused(),
    });

  return {
    tryHandleOverlayShortcutLocalFallback: (inputEvent) =>
      runOverlayShortcutLocalFallback(
        inputEvent,
        input.getConfiguredShortcuts(),
        shortcutMatchesInputForLocalFallback,
        handlers.fallbackHandlers,
      ),
    registerOverlayShortcuts: () => {
      input.setShortcutsRegistered(
        registerOverlayShortcuts(input.getConfiguredShortcuts(), handlers.overlayHandlers),
      );
    },
    unregisterOverlayShortcuts: () => {
      input.setShortcutsRegistered(
        unregisterOverlayShortcutsRuntime(
          input.getShortcutsRegistered(),
          getShortcutLifecycleDeps(),
        ),
      );
    },
    syncOverlayShortcuts: () => {
      input.setShortcutsRegistered(
        syncOverlayShortcutsRuntime(
          shouldOverlayShortcutsBeActive(),
          input.getShortcutsRegistered(),
          getShortcutLifecycleDeps(),
        ),
      );
    },
    refreshOverlayShortcuts: () => {
      input.setShortcutsRegistered(
        refreshOverlayShortcutsRuntime(
          shouldOverlayShortcutsBeActive(),
          input.getShortcutsRegistered(),
          getShortcutLifecycleDeps(),
        ),
      );
    },
  };
}
