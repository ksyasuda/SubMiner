import { OverlayShortcutHandlers, registerOverlayShortcutsService, unregisterOverlayShortcutsService } from "./overlay-shortcut-service";
import { ConfiguredShortcuts } from "../utils/shortcut-config";

export interface OverlayShortcutLifecycleDeps {
  getConfiguredShortcuts: () => ConfiguredShortcuts;
  getOverlayHandlers: () => OverlayShortcutHandlers;
  cancelPendingMultiCopy: () => void;
  cancelPendingMineSentenceMultiple: () => void;
}

export function registerOverlayShortcutsRuntimeService(
  deps: OverlayShortcutLifecycleDeps,
): boolean {
  return registerOverlayShortcutsService(
    deps.getConfiguredShortcuts(),
    deps.getOverlayHandlers(),
  );
}

export function unregisterOverlayShortcutsRuntimeService(
  shortcutsRegistered: boolean,
  deps: OverlayShortcutLifecycleDeps,
): boolean {
  if (!shortcutsRegistered) return shortcutsRegistered;
  deps.cancelPendingMultiCopy();
  deps.cancelPendingMineSentenceMultiple();
  unregisterOverlayShortcutsService(deps.getConfiguredShortcuts());
  return false;
}

export function syncOverlayShortcutsRuntimeService(
  shouldBeActive: boolean,
  shortcutsRegistered: boolean,
  deps: OverlayShortcutLifecycleDeps,
): boolean {
  if (shouldBeActive) {
    return registerOverlayShortcutsRuntimeService(deps);
  }
  return unregisterOverlayShortcutsRuntimeService(shortcutsRegistered, deps);
}

export function refreshOverlayShortcutsRuntimeService(
  shouldBeActive: boolean,
  shortcutsRegistered: boolean,
  deps: OverlayShortcutLifecycleDeps,
): boolean {
  const cleared = unregisterOverlayShortcutsRuntimeService(
    shortcutsRegistered,
    deps,
  );
  return syncOverlayShortcutsRuntimeService(shouldBeActive, cleared, deps);
}
