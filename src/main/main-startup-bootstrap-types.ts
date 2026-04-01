import type { CliArgs } from '../cli/args';
import type { ResolvedConfig, SecondarySubMode, SubtitleData } from '../types';
import { RuntimeOptionsManager } from '../runtime-options';
import { SubtitleTimingTracker } from '../subtitle-timing-tracker';

export type StartupBootstrapMpvClientLike = {
  connected: boolean;
  connect: () => void;
  setSocketPath: (socketPath: string) => void;
  currentSubStart?: number | null;
  currentSubEnd?: number | null;
};

export type StartupBootstrapAppStateLike = {
  subtitlePosition: unknown | null;
  keybindings: unknown[];
  mpvSocketPath: string;
  texthookerPort: number;
  mpvClient: StartupBootstrapMpvClientLike | null;
  runtimeOptionsManager: RuntimeOptionsManager | null;
  subtitleTimingTracker: SubtitleTimingTracker | null;
  currentSubtitleData: SubtitleData | null;
  currentSubText: string | null;
  initialArgs: CliArgs | null | undefined;
  backgroundMode: boolean;
  texthookerOnlyMode: boolean;
  overlayRuntimeInitialized: boolean;
  firstRunSetupCompleted: boolean;
  secondarySubMode: SecondarySubMode;
  ankiIntegration: unknown | null;
  immersionTracker: unknown | null;
};

export type StartupBootstrapSubtitleWebsocketLike = {
  start: (
    port: number,
    getPayload: () => SubtitleData | null,
    getFrequencyOptions: () => {
      enabled: boolean;
      topX: number;
      mode: ResolvedConfig['subtitleStyle']['frequencyDictionary']['mode'];
    },
  ) => void;
};

export type StartupBootstrapOverlayUiLike = {
  broadcastRuntimeOptionsChanged: () => void;
  ensureOverlayWindowsReadyForVisibilityActions: () => void;
  ensureTray: () => void;
  initializeOverlayRuntime: () => void;
  openRuntimeOptionsPalette: () => void;
  setVisibleOverlayVisible: (visible: boolean) => void;
  toggleVisibleOverlay: () => void;
};
