import type { BrowserWindow, Extension, Session } from 'electron';

import type {
  Keybinding,
  MpvSubtitleRenderMetrics,
  SecondarySubMode,
  SubtitleData,
  SubtitlePosition,
  KikuFieldGroupingChoice,
  JlptLevel,
  FrequencyDictionaryLookup,
} from '../types';
import type { CliArgs } from '../cli/args';
import type { SubtitleTimingTracker } from '../subtitle-timing-tracker';
import type { AnkiIntegration } from '../anki-integration';
import type { ImmersionTrackerService } from '../core/services/immersion-tracker-service';
import type { MpvIpcClient } from '../core/services/mpv';
import type { JellyfinRemoteSessionService } from '../core/services/jellyfin-remote';
import type { createDiscordPresenceService } from '../core/services/discord-presence';
import { DEFAULT_MPV_SUBTITLE_RENDER_METRICS } from '../core/services/mpv-render-metrics';
import type { RuntimeOptionsManager } from '../runtime-options';
import type { MecabTokenizer } from '../mecab-tokenizer';
import type { BaseWindowTracker } from '../window-trackers';
import type { AnilistMediaGuess } from '../core/services/anilist/anilist-updater';

export interface AnilistSecretResolutionState {
  status: 'not_checked' | 'resolved' | 'error';
  source: 'none' | 'literal' | 'stored';
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

export interface AnilistMediaGuessRuntimeState {
  mediaKey: string | null;
  mediaDurationSec: number | null;
  mediaGuess: AnilistMediaGuess | null;
  mediaGuessPromise: Promise<AnilistMediaGuess | null> | null;
  lastDurationProbeAtMs: number;
}

export interface AnilistUpdateInFlightState {
  inFlight: boolean;
}

export function createInitialAnilistSecretResolutionState(): AnilistSecretResolutionState {
  return {
    status: 'not_checked',
    source: 'none',
    message: null,
    resolvedAt: null,
    errorAt: null,
  };
}

export function createInitialAnilistRetryQueueState(): AnilistRetryQueueState {
  return {
    pending: 0,
    ready: 0,
    deadLetter: 0,
    lastAttemptAt: null,
    lastError: null,
  };
}

export function createInitialAnilistMediaGuessRuntimeState(): AnilistMediaGuessRuntimeState {
  return {
    mediaKey: null,
    mediaDurationSec: null,
    mediaGuess: null,
    mediaGuessPromise: null,
    lastDurationProbeAtMs: 0,
  };
}

export function createInitialAnilistUpdateInFlightState(): AnilistUpdateInFlightState {
  return {
    inFlight: false,
  };
}

export function transitionAnilistClientSecretState(
  _current: AnilistSecretResolutionState,
  next: AnilistSecretResolutionState,
): AnilistSecretResolutionState {
  return next;
}

export function transitionAnilistRetryQueueState(
  _current: AnilistRetryQueueState,
  next: AnilistRetryQueueState,
): AnilistRetryQueueState {
  return next;
}

export function transitionAnilistRetryQueueLastAttemptAt(
  current: AnilistRetryQueueState,
  lastAttemptAt: number | null,
): AnilistRetryQueueState {
  return {
    ...current,
    lastAttemptAt,
  };
}

export function transitionAnilistRetryQueueLastError(
  current: AnilistRetryQueueState,
  lastError: string | null,
): AnilistRetryQueueState {
  return {
    ...current,
    lastError,
  };
}

export function transitionAnilistMediaGuessRuntimeState(
  current: AnilistMediaGuessRuntimeState,
  partial: Partial<AnilistMediaGuessRuntimeState>,
): AnilistMediaGuessRuntimeState {
  return {
    ...current,
    ...partial,
  };
}

export function transitionAnilistUpdateInFlightState(
  current: AnilistUpdateInFlightState,
  inFlight: boolean,
): AnilistUpdateInFlightState {
  return {
    ...current,
    inFlight,
  };
}

export interface AppState {
  yomitanExt: Extension | null;
  yomitanSession: Session | null;
  yomitanSettingsWindow: BrowserWindow | null;
  yomitanParserWindow: BrowserWindow | null;
  anilistSetupWindow: BrowserWindow | null;
  jellyfinSetupWindow: BrowserWindow | null;
  firstRunSetupWindow: BrowserWindow | null;
  yomitanParserReadyPromise: Promise<void> | null;
  yomitanParserInitPromise: Promise<boolean> | null;
  mpvClient: MpvIpcClient | null;
  jellyfinRemoteSession: JellyfinRemoteSessionService | null;
  discordPresenceService: ReturnType<typeof createDiscordPresenceService> | null;
  reconnectTimer: ReturnType<typeof setTimeout> | null;
  currentSubText: string;
  currentSubAssText: string;
  currentSubtitleData: SubtitleData | null;
  windowTracker: BaseWindowTracker | null;
  subtitlePosition: SubtitlePosition | null;
  currentMediaPath: string | null;
  currentMediaTitle: string | null;
  playbackPaused: boolean | null;
  pendingSubtitlePosition: SubtitlePosition | null;
  anilistClientSecretState: AnilistSecretResolutionState;
  mecabTokenizer: MecabTokenizer | null;
  keybindings: Keybinding[];
  subtitleTimingTracker: SubtitleTimingTracker | null;
  immersionTracker: ImmersionTrackerService | null;
  ankiIntegration: AnkiIntegration | null;
  secondarySubMode: SecondarySubMode;
  lastSecondarySubToggleAtMs: number;
  previousSecondarySubVisibility: boolean | null;
  overlaySavedMpvSubVisibility: boolean | null;
  overlayMpvSubVisibilityRevision: number;
  mpvSubtitleRenderMetrics: MpvSubtitleRenderMetrics;
  shortcutsRegistered: boolean;
  overlayRuntimeInitialized: boolean;
  fieldGroupingResolver: ((choice: KikuFieldGroupingChoice) => void) | null;
  fieldGroupingResolverSequence: number;
  runtimeOptionsManager: RuntimeOptionsManager | null;
  trackerNotReadyWarningShown: boolean;
  overlayDebugVisualizationEnabled: boolean;
  statsOverlayVisible: boolean;
  subsyncInProgress: boolean;
  initialArgs: CliArgs | null;
  mpvSocketPath: string;
  texthookerPort: number;
  backendOverride: string | null;
  autoStartOverlay: boolean;
  texthookerOnlyMode: boolean;
  backgroundMode: boolean;
  jlptLevelLookup: (term: string) => JlptLevel | null;
  frequencyRankLookup: FrequencyDictionaryLookup;
  anilistSetupPageOpened: boolean;
  anilistRetryQueueState: AnilistRetryQueueState;
  firstRunSetupCompleted: boolean;
  statsServer: { close: () => void } | null;
  statsStartupInProgress: boolean;
}

export interface AppStateInitialValues {
  mpvSocketPath: string;
  texthookerPort: number;
  backendOverride?: string | null;
  autoStartOverlay?: boolean;
  texthookerOnlyMode?: boolean;
  backgroundMode?: boolean;
}

export interface StartupState {
  initialArgs: Exclude<AppState['initialArgs'], null>;
  mpvSocketPath: AppState['mpvSocketPath'];
  texthookerPort: AppState['texthookerPort'];
  backendOverride: AppState['backendOverride'];
  autoStartOverlay: AppState['autoStartOverlay'];
  texthookerOnlyMode: AppState['texthookerOnlyMode'];
  backgroundMode: AppState['backgroundMode'];
}

export function createAppState(values: AppStateInitialValues): AppState {
  return {
    yomitanExt: null,
    yomitanSession: null,
    yomitanSettingsWindow: null,
    yomitanParserWindow: null,
    anilistSetupWindow: null,
    jellyfinSetupWindow: null,
    firstRunSetupWindow: null,
    yomitanParserReadyPromise: null,
    yomitanParserInitPromise: null,
    mpvClient: null,
    jellyfinRemoteSession: null,
    discordPresenceService: null,
    reconnectTimer: null,
    currentSubText: '',
    currentSubAssText: '',
    currentSubtitleData: null,
    windowTracker: null,
    subtitlePosition: null,
    currentMediaPath: null,
    currentMediaTitle: null,
    playbackPaused: null,
    pendingSubtitlePosition: null,
    anilistClientSecretState: createInitialAnilistSecretResolutionState(),
    mecabTokenizer: null,
    keybindings: [],
    subtitleTimingTracker: null,
    immersionTracker: null,
    ankiIntegration: null,
    secondarySubMode: 'hover',
    lastSecondarySubToggleAtMs: 0,
    previousSecondarySubVisibility: null,
    overlaySavedMpvSubVisibility: null,
    overlayMpvSubVisibilityRevision: 0,
    mpvSubtitleRenderMetrics: {
      ...DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
    },
    runtimeOptionsManager: null,
    trackerNotReadyWarningShown: false,
    overlayDebugVisualizationEnabled: false,
    statsOverlayVisible: false,
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
    backgroundMode: values.backgroundMode ?? false,
    jlptLevelLookup: () => null,
    frequencyRankLookup: () => null,
    anilistSetupPageOpened: false,
    anilistRetryQueueState: createInitialAnilistRetryQueueState(),
    firstRunSetupCompleted: false,
    statsServer: null,
    statsStartupInProgress: false,
  };
}

export function applyStartupState(appState: AppState, startupState: StartupState): void {
  appState.initialArgs = startupState.initialArgs;
  appState.mpvSocketPath = startupState.mpvSocketPath;
  appState.texthookerPort = startupState.texthookerPort;
  appState.backendOverride = startupState.backendOverride;
  appState.autoStartOverlay = startupState.autoStartOverlay;
  appState.texthookerOnlyMode = startupState.texthookerOnlyMode;
  appState.backgroundMode = startupState.backgroundMode;
}
