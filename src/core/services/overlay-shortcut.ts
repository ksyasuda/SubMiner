import { ConfiguredShortcuts } from '../utils/shortcut-config';

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
  openJimaku: () => void;
}

export interface OverlayShortcutLifecycleDeps {
  getConfiguredShortcuts: () => ConfiguredShortcuts;
  getOverlayHandlers: () => OverlayShortcutHandlers;
  cancelPendingMultiCopy: () => void;
  cancelPendingMineSentenceMultiple: () => void;
}

export function shouldActivateOverlayShortcuts(args: {
  overlayRuntimeInitialized: boolean;
  isMacOSPlatform: boolean;
  trackedMpvWindowFocused: boolean;
}): boolean {
  if (!args.overlayRuntimeInitialized) {
    return false;
  }
  if (!args.isMacOSPlatform) {
    return true;
  }
  return args.trackedMpvWindowFocused;
}

export function registerOverlayShortcuts(
  _shortcuts: ConfiguredShortcuts,
  _handlers: OverlayShortcutHandlers,
): boolean {
  return false;
}

export function unregisterOverlayShortcuts(_shortcuts: ConfiguredShortcuts): void {}

export function registerOverlayShortcutsRuntime(deps: OverlayShortcutLifecycleDeps): boolean {
  return registerOverlayShortcuts(deps.getConfiguredShortcuts(), deps.getOverlayHandlers());
}

export function unregisterOverlayShortcutsRuntime(
  shortcutsRegistered: boolean,
  deps: OverlayShortcutLifecycleDeps,
): boolean {
  if (!shortcutsRegistered) return shortcutsRegistered;
  deps.cancelPendingMultiCopy();
  deps.cancelPendingMineSentenceMultiple();
  unregisterOverlayShortcuts(deps.getConfiguredShortcuts());
  return false;
}

export function syncOverlayShortcutsRuntime(
  shouldBeActive: boolean,
  shortcutsRegistered: boolean,
  deps: OverlayShortcutLifecycleDeps,
): boolean {
  if (shouldBeActive) {
    return registerOverlayShortcutsRuntime(deps);
  }
  return unregisterOverlayShortcutsRuntime(shortcutsRegistered, deps);
}

export function refreshOverlayShortcutsRuntime(
  shouldBeActive: boolean,
  shortcutsRegistered: boolean,
  deps: OverlayShortcutLifecycleDeps,
): boolean {
  const cleared = unregisterOverlayShortcutsRuntime(shortcutsRegistered, deps);
  return syncOverlayShortcutsRuntime(shouldBeActive, cleared, deps);
}
