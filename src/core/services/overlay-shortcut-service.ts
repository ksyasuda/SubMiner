import { globalShortcut } from "electron";
import { ConfiguredShortcuts } from "../utils/shortcut-config";
import { isGlobalShortcutRegisteredSafe } from "./shortcut-fallback-service";

export interface OverlayShortcutHandlers {
  copySubtitle: () => void;
  copySubtitleMultiple: (timeoutMs: number) => void;
  updateLastCardFromClipboard: () => void;
  triggerFieldGrouping: () => void;
  triggerSubsync: () => void;
  mineSentence: () => void;
  mineSentenceMultiple: (timeoutMs: number) => void;
  toggleSecondarySub: () => void;
  markAudioCard: () => void;
  openRuntimeOptions: () => void;
}

export function registerOverlayShortcutsService(
  shortcuts: ConfiguredShortcuts,
  handlers: OverlayShortcutHandlers,
): boolean {
  let registeredAny = false;
  const registerOverlayShortcut = (
    accelerator: string,
    handler: () => void,
    label: string,
  ): void => {
    if (isGlobalShortcutRegisteredSafe(accelerator)) {
      registeredAny = true;
      return;
    }
    const ok = globalShortcut.register(accelerator, handler);
    if (!ok) {
      console.warn(
        `Failed to register overlay shortcut ${label}: ${accelerator}`,
      );
      return;
    }
    registeredAny = true;
  };

  if (shortcuts.copySubtitleMultiple) {
    registerOverlayShortcut(
      shortcuts.copySubtitleMultiple,
      () => handlers.copySubtitleMultiple(shortcuts.multiCopyTimeoutMs),
      "copySubtitleMultiple",
    );
  }

  if (shortcuts.copySubtitle) {
    registerOverlayShortcut(
      shortcuts.copySubtitle,
      () => handlers.copySubtitle(),
      "copySubtitle",
    );
  }

  if (shortcuts.triggerFieldGrouping) {
    registerOverlayShortcut(
      shortcuts.triggerFieldGrouping,
      () => handlers.triggerFieldGrouping(),
      "triggerFieldGrouping",
    );
  }

  if (shortcuts.triggerSubsync) {
    registerOverlayShortcut(
      shortcuts.triggerSubsync,
      () => handlers.triggerSubsync(),
      "triggerSubsync",
    );
  }

  if (shortcuts.mineSentence) {
    registerOverlayShortcut(
      shortcuts.mineSentence,
      () => handlers.mineSentence(),
      "mineSentence",
    );
  }

  if (shortcuts.mineSentenceMultiple) {
    registerOverlayShortcut(
      shortcuts.mineSentenceMultiple,
      () => handlers.mineSentenceMultiple(shortcuts.multiCopyTimeoutMs),
      "mineSentenceMultiple",
    );
  }

  if (shortcuts.toggleSecondarySub) {
    registerOverlayShortcut(
      shortcuts.toggleSecondarySub,
      () => handlers.toggleSecondarySub(),
      "toggleSecondarySub",
    );
  }

  if (shortcuts.updateLastCardFromClipboard) {
    registerOverlayShortcut(
      shortcuts.updateLastCardFromClipboard,
      () => handlers.updateLastCardFromClipboard(),
      "updateLastCardFromClipboard",
    );
  }

  if (shortcuts.markAudioCard) {
    registerOverlayShortcut(
      shortcuts.markAudioCard,
      () => handlers.markAudioCard(),
      "markAudioCard",
    );
  }

  if (shortcuts.openRuntimeOptions) {
    registerOverlayShortcut(
      shortcuts.openRuntimeOptions,
      () => handlers.openRuntimeOptions(),
      "openRuntimeOptions",
    );
  }

  return registeredAny;
}

export function unregisterOverlayShortcutsService(
  shortcuts: ConfiguredShortcuts,
): void {
  if (shortcuts.copySubtitle) {
    globalShortcut.unregister(shortcuts.copySubtitle);
  }
  if (shortcuts.copySubtitleMultiple) {
    globalShortcut.unregister(shortcuts.copySubtitleMultiple);
  }
  if (shortcuts.updateLastCardFromClipboard) {
    globalShortcut.unregister(shortcuts.updateLastCardFromClipboard);
  }
  if (shortcuts.triggerFieldGrouping) {
    globalShortcut.unregister(shortcuts.triggerFieldGrouping);
  }
  if (shortcuts.triggerSubsync) {
    globalShortcut.unregister(shortcuts.triggerSubsync);
  }
  if (shortcuts.mineSentence) {
    globalShortcut.unregister(shortcuts.mineSentence);
  }
  if (shortcuts.mineSentenceMultiple) {
    globalShortcut.unregister(shortcuts.mineSentenceMultiple);
  }
  if (shortcuts.toggleSecondarySub) {
    globalShortcut.unregister(shortcuts.toggleSecondarySub);
  }
  if (shortcuts.markAudioCard) {
    globalShortcut.unregister(shortcuts.markAudioCard);
  }
  if (shortcuts.openRuntimeOptions) {
    globalShortcut.unregister(shortcuts.openRuntimeOptions);
  }
}
