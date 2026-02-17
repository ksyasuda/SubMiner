import type { BrowserWindow, Extension } from "electron";

import type {
  Keybinding,
  MpvSubtitleRenderMetrics,
  SecondarySubMode,
  SubtitlePosition,
  KikuFieldGroupingChoice,
  JlptLevel,
  FrequencyDictionaryLookup,
} from "../types";
import type { CliArgs } from "../cli/args";
import type { SubtitleTimingTracker } from "../subtitle-timing-tracker";
import type { AnkiIntegration } from "../anki-integration";
import type { MpvIpcClient } from "../core/services";
import { DEFAULT_MPV_SUBTITLE_RENDER_METRICS } from "../core/services";
import type { RuntimeOptionsManager } from "../runtime-options";
import type { MecabTokenizer } from "../mecab-tokenizer";
import type { BaseWindowTracker } from "../window-trackers";

export interface AnilistSecretResolutionState {
  status: "not_checked" | "resolved" | "error";
  source: "none" | "literal" | "stored";
  message: string | null;
  resolvedAt: number | null;
  errorAt: number | null;
}

export interface AnilistRetryQueueState {
  pending: number;
  ready: number;
  deadLetter: number;
  lastAttemptAt: number | null;
  lastError: string | null;
}

export interface AppState {
  yomitanExt: Extension | null;
  yomitanSettingsWindow: BrowserWindow | null;
  yomitanParserWindow: BrowserWindow | null;
  anilistSetupWindow: BrowserWindow | null;
  yomitanParserReadyPromise: Promise<void> | null;
  yomitanParserInitPromise: Promise<boolean> | null;
  mpvClient: MpvIpcClient | null;
  reconnectTimer: ReturnType<typeof setTimeout> | null;
  currentSubText: string;
  currentSubAssText: string;
  windowTracker: BaseWindowTracker | null;
  subtitlePosition: SubtitlePosition | null;
  currentMediaPath: string | null;
  currentMediaTitle: string | null;
  pendingSubtitlePosition: SubtitlePosition | null;
  anilistClientSecretState: AnilistSecretResolutionState;
  mecabTokenizer: MecabTokenizer | null;
  keybindings: Keybinding[];
  subtitleTimingTracker: SubtitleTimingTracker | null;
  ankiIntegration: AnkiIntegration | null;
  secondarySubMode: SecondarySubMode;
  lastSecondarySubToggleAtMs: number;
  previousSecondarySubVisibility: boolean | null;
  mpvSubtitleRenderMetrics: MpvSubtitleRenderMetrics;
  shortcutsRegistered: boolean;
  overlayRuntimeInitialized: boolean;
  fieldGroupingResolver: ((choice: KikuFieldGroupingChoice) => void) | null;
  fieldGroupingResolverSequence: number;
  runtimeOptionsManager: RuntimeOptionsManager | null;
  trackerNotReadyWarningShown: boolean;
  overlayDebugVisualizationEnabled: boolean;
  subsyncInProgress: boolean;
  initialArgs: CliArgs | null;
  mpvSocketPath: string;
  texthookerPort: number;
  backendOverride: string | null;
  autoStartOverlay: boolean;
  texthookerOnlyMode: boolean;
  jlptLevelLookup: (term: string) => JlptLevel | null;
  frequencyRankLookup: FrequencyDictionaryLookup;
  anilistSetupPageOpened: boolean;
  anilistRetryQueueState: AnilistRetryQueueState;
}

export interface AppStateInitialValues {
  mpvSocketPath: string;
  texthookerPort: number;
  backendOverride?: string | null;
  autoStartOverlay?: boolean;
  texthookerOnlyMode?: boolean;
}

export interface StartupState {
  initialArgs: Exclude<AppState["initialArgs"], null>;
  mpvSocketPath: AppState["mpvSocketPath"];
  texthookerPort: AppState["texthookerPort"];
  backendOverride: AppState["backendOverride"];
  autoStartOverlay: AppState["autoStartOverlay"];
  texthookerOnlyMode: AppState["texthookerOnlyMode"];
}

export function createAppState(values: AppStateInitialValues): AppState {
  return {
    yomitanExt: null,
    yomitanSettingsWindow: null,
    yomitanParserWindow: null,
    anilistSetupWindow: null,
    yomitanParserReadyPromise: null,
    yomitanParserInitPromise: null,
    mpvClient: null,
    reconnectTimer: null,
    currentSubText: "",
    currentSubAssText: "",
    windowTracker: null,
    subtitlePosition: null,
    currentMediaPath: null,
    currentMediaTitle: null,
    pendingSubtitlePosition: null,
    anilistClientSecretState: {
      status: "not_checked",
      source: "none",
      message: null,
      resolvedAt: null,
      errorAt: null,
    },
    mecabTokenizer: null,
    keybindings: [],
    subtitleTimingTracker: null,
    ankiIntegration: null,
    secondarySubMode: "hover",
    lastSecondarySubToggleAtMs: 0,
    previousSecondarySubVisibility: null,
    mpvSubtitleRenderMetrics: {
      ...DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
    },
    runtimeOptionsManager: null,
    trackerNotReadyWarningShown: false,
    overlayDebugVisualizationEnabled: false,
    shortcutsRegistered: false,
    overlayRuntimeInitialized: false,
    fieldGroupingResolver: null,
    fieldGroupingResolverSequence: 0,
    subsyncInProgress: false,
    initialArgs: null,
    mpvSocketPath: values.mpvSocketPath,
    texthookerPort: values.texthookerPort,
    backendOverride: values.backendOverride ?? null,
    autoStartOverlay: values.autoStartOverlay ?? false,
    texthookerOnlyMode: values.texthookerOnlyMode ?? false,
    jlptLevelLookup: () => null,
    frequencyRankLookup: () => null,
    anilistSetupPageOpened: false,
    anilistRetryQueueState: {
      pending: 0,
      ready: 0,
      deadLetter: 0,
      lastAttemptAt: null,
      lastError: null,
    },
  };
}

export function applyStartupState(
  appState: AppState,
  startupState: StartupState,
): void {
  appState.initialArgs = startupState.initialArgs;
  appState.mpvSocketPath = startupState.mpvSocketPath;
  appState.texthookerPort = startupState.texthookerPort;
  appState.backendOverride = startupState.backendOverride;
  appState.autoStartOverlay = startupState.autoStartOverlay;
  appState.texthookerOnlyMode = startupState.texthookerOnlyMode;
}
