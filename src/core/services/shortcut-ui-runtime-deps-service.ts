import { Extension } from "electron";
import { SecondarySubMode } from "../../types";
import { ConfiguredShortcuts } from "../utils/shortcut-config";
import { CycleSecondarySubModeDeps } from "./secondary-subtitle-service";
import { OverlayShortcutFallbackHandlers, runOverlayShortcutLocalFallback } from "./overlay-shortcut-fallback-runner";
import { OpenYomitanSettingsWindowOptions } from "./yomitan-settings-service";
import { RegisterGlobalShortcutsServiceOptions } from "./shortcut-service";

export interface ShortcutUiRuntimeDepsOptions {
  yomitanExt: Extension | null;
  getYomitanSettingsWindow: OpenYomitanSettingsWindowOptions["getExistingWindow"];
  setYomitanSettingsWindow: OpenYomitanSettingsWindowOptions["setWindow"];

  shortcuts: RegisterGlobalShortcutsServiceOptions["shortcuts"];
  onToggleVisibleOverlay: () => void;
  onToggleInvisibleOverlay: () => void;
  onOpenYomitanSettings: () => void;
  isDev: boolean;
  getMainWindow: RegisterGlobalShortcutsServiceOptions["getMainWindow"];

  getSecondarySubMode: () => SecondarySubMode;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  getLastSecondarySubToggleAtMs: () => number;
  setLastSecondarySubToggleAtMs: (timestampMs: number) => void;
  broadcastSecondarySubMode: (mode: SecondarySubMode) => void;
  showMpvOsd: (text: string) => void;

  getConfiguredShortcuts: () => ConfiguredShortcuts;
  getOverlayShortcutFallbackHandlers: () => OverlayShortcutFallbackHandlers;
  shortcutMatcher: (
    input: Electron.Input,
    accelerator: string,
    allowWhenRegistered?: boolean,
  ) => boolean;
}

export function createYomitanSettingsWindowDepsRuntimeService(
  options: ShortcutUiRuntimeDepsOptions,
): OpenYomitanSettingsWindowOptions {
  return {
    yomitanExt: options.yomitanExt,
    getExistingWindow: options.getYomitanSettingsWindow,
    setWindow: options.setYomitanSettingsWindow,
  };
}

export function createGlobalShortcutRegistrationDepsRuntimeService(
  options: ShortcutUiRuntimeDepsOptions,
): RegisterGlobalShortcutsServiceOptions {
  return {
    shortcuts: options.shortcuts,
    onToggleVisibleOverlay: options.onToggleVisibleOverlay,
    onToggleInvisibleOverlay: options.onToggleInvisibleOverlay,
    onOpenYomitanSettings: options.onOpenYomitanSettings,
    isDev: options.isDev,
    getMainWindow: options.getMainWindow,
  };
}

export function createSecondarySubtitleCycleDepsRuntimeService(
  options: ShortcutUiRuntimeDepsOptions,
): CycleSecondarySubModeDeps {
  return {
    getSecondarySubMode: options.getSecondarySubMode,
    setSecondarySubMode: options.setSecondarySubMode,
    getLastSecondarySubToggleAtMs: options.getLastSecondarySubToggleAtMs,
    setLastSecondarySubToggleAtMs: options.setLastSecondarySubToggleAtMs,
    broadcastSecondarySubMode: options.broadcastSecondarySubMode,
    showMpvOsd: options.showMpvOsd,
  };
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
