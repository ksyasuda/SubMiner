import type { BrowserWindow } from 'electron';
import type { BaseWindowTracker } from '../../window-trackers';
import type {
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  WindowGeometry,
} from '../../types';

type InitializeOverlayRuntimeCore = (options: {
  backendOverride: string | null;
  getInitialInvisibleOverlayVisibility: () => boolean;
  createMainWindow: () => void;
  createInvisibleWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  updateInvisibleOverlayBounds: (geometry: WindowGeometry) => void;
  isVisibleOverlayVisible: () => boolean;
  isInvisibleOverlayVisible: () => boolean;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  setWindowTracker: (tracker: BaseWindowTracker | null) => void;
  getMpvSocketPath: () => string;
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  getSubtitleTimingTracker: () => unknown | null;
  getMpvClient: () => { send?: (payload: { command: string[] }) => void } | null;
  getRuntimeOptionsManager: () => {
    getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
  } | null;
  setAnkiIntegration: (integration: unknown | null) => void;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  getKnownWordCacheStatePath: () => string;
}) => { invisibleOverlayVisible: boolean };

export function createInitializeOverlayRuntimeHandler(deps: {
  isOverlayRuntimeInitialized: () => boolean;
  initializeOverlayRuntimeCore: InitializeOverlayRuntimeCore;
  buildOptions: () => Parameters<InitializeOverlayRuntimeCore>[0];
  setInvisibleOverlayVisible: (visible: boolean) => void;
  setOverlayRuntimeInitialized: (initialized: boolean) => void;
  startBackgroundWarmups: () => void;
}) {
  return (): void => {
    if (deps.isOverlayRuntimeInitialized()) return;
    const result = deps.initializeOverlayRuntimeCore(deps.buildOptions());
    deps.setInvisibleOverlayVisible(result.invisibleOverlayVisible);
    deps.setOverlayRuntimeInitialized(true);
    deps.startBackgroundWarmups();
  };
}
