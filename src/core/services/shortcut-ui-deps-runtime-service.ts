import { ConfiguredShortcuts } from "../utils/shortcut-config";
import { OverlayShortcutFallbackHandlers, runOverlayShortcutLocalFallback } from "./overlay-shortcut-fallback-runner";

export interface ShortcutUiRuntimeDepsOptions {
  getConfiguredShortcuts: () => ConfiguredShortcuts;
  getOverlayShortcutFallbackHandlers: () => OverlayShortcutFallbackHandlers;
  shortcutMatcher: (
    input: Electron.Input,
    accelerator: string,
    allowWhenRegistered?: boolean,
  ) => boolean;
}

export function runOverlayShortcutLocalFallbackRuntimeService(
  input: Electron.Input,
  options: ShortcutUiRuntimeDepsOptions,
): boolean {
  return runOverlayShortcutLocalFallback(
    input,
    options.getConfiguredShortcuts(),
    options.shortcutMatcher,
    options.getOverlayShortcutFallbackHandlers(),
  );
}
