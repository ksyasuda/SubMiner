import type { BrowserWindow } from 'electron';
import { BaseWindowTracker, createWindowTracker } from '../../window-trackers';
import { mergeAiConfig } from '../../ai/config';
import {
  AiConfig,
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  OverlayNotificationPayload,
  WindowGeometry,
} from '../../types';

type AnkiIntegrationLike = {
  start: () => void;
};

type CreateAnkiIntegrationArgs = {
  config: AnkiConnectConfig;
  aiConfig: AiConfig;
  subtitleTimingTracker: unknown;
  mpvClient: { send?: (payload: { command: string[] }) => void };
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  knownWordCacheStatePath: string;
};

export type OverlayWindowTrackerOptions = {
  backendOverride: string | null;
  getMpvSocketPath: () => string;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  isVisibleOverlayVisible: () => boolean;
  updateVisibleOverlayVisibility: () => void;
  refreshCurrentSubtitle?: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  setWindowTracker: (tracker: BaseWindowTracker | null) => void;
  createWindowTracker?: (
    override?: string | null,
    targetMpvSocketPath?: string | null,
  ) => BaseWindowTracker | null;
  bindOverlayOwner?: () => void;
  releaseOverlayOwner?: () => void;
};

function createDefaultAnkiIntegration(args: CreateAnkiIntegrationArgs): AnkiIntegrationLike {
  const { AnkiIntegration } =
    require('../../anki-integration') as typeof import('../../anki-integration');
  return new AnkiIntegration(
    args.config,
    args.subtitleTimingTracker as never,
    args.mpvClient as never,
    (text: string) => {
      if (args.mpvClient && typeof args.mpvClient.send === 'function') {
        args.mpvClient.send({
          command: ['show-text', text, '3000'],
        });
      }
    },
    args.showDesktopNotification,
    args.createFieldGroupingCallback(),
    args.knownWordCacheStatePath,
    args.aiConfig,
    undefined,
    args.showOverlayNotification,
  );
}

export function startOverlayWindowTracker(
  options: OverlayWindowTrackerOptions,
): BaseWindowTracker | null {
  const createWindowTrackerHandler = options.createWindowTracker ?? createWindowTracker;
  const windowTracker = createWindowTrackerHandler(
    options.backendOverride,
    options.getMpvSocketPath(),
  );
  options.setWindowTracker(windowTracker);
  if (!windowTracker) {
    return null;
  }

  windowTracker.onGeometryChange = (geometry: WindowGeometry) => {
    options.updateVisibleOverlayBounds(geometry);
  };
  windowTracker.onWindowFound = (geometry: WindowGeometry) => {
    options.updateVisibleOverlayBounds(geometry);
    options.bindOverlayOwner?.();
    if (options.isVisibleOverlayVisible()) {
      options.updateVisibleOverlayVisibility();
      options.refreshCurrentSubtitle?.();
    }
  };
  windowTracker.onWindowLost = () => {
    options.releaseOverlayOwner?.();
    if (windowTracker.isTargetWindowMinimized()) {
      for (const window of options.getOverlayWindows()) {
        window.hide();
      }
      options.syncOverlayShortcuts();
      return;
    }
    options.updateVisibleOverlayVisibility();
  };
  windowTracker.onWindowFocusChange = () => {
    if (options.isVisibleOverlayVisible()) {
      options.updateVisibleOverlayVisibility();
    }
    options.syncOverlayShortcuts();
  };
  windowTracker.start();
  return windowTracker;
}

export function initializeOverlayRuntime(
  options: OverlayWindowTrackerOptions & {
    getMpvSocketPath: () => string;
    getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig; ai?: AiConfig };
    getSubtitleTimingTracker: () => unknown | null;
    getMpvClient: () => {
      send?: (payload: { command: string[] }) => void;
    } | null;
    getRuntimeOptionsManager: () => {
      getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
    } | null;
    getAnkiIntegration?: () => unknown | null;
    setAnkiIntegration: (integration: unknown | null) => void;
    showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
    showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
    createFieldGroupingCallback: () => (
      data: KikuFieldGroupingRequestData,
    ) => Promise<KikuFieldGroupingChoice>;
    getKnownWordCacheStatePath: () => string;
    shouldStartAnkiIntegration?: () => boolean;
    createAnkiIntegration?: (args: CreateAnkiIntegrationArgs) => AnkiIntegrationLike;
    backendOverride: string | null;
    createMainWindow: () => void;
    registerGlobalShortcuts: () => void;
  },
): void {
  options.createMainWindow();
  options.registerGlobalShortcuts();

  startOverlayWindowTracker(options);

  initializeOverlayAnkiIntegration(options);

  options.updateVisibleOverlayVisibility();
}

export function initializeOverlayAnkiIntegration(options: {
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig; ai?: AiConfig };
  getSubtitleTimingTracker: () => unknown | null;
  getMpvClient: () => {
    send?: (payload: { command: string[] }) => void;
  } | null;
  getRuntimeOptionsManager: () => {
    getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
  } | null;
  getAnkiIntegration?: () => unknown | null;
  setAnkiIntegration: (integration: unknown | null) => void;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  getKnownWordCacheStatePath: () => string;
  shouldStartAnkiIntegration?: () => boolean;
  createAnkiIntegration?: (args: CreateAnkiIntegrationArgs) => AnkiIntegrationLike;
}): boolean {
  if (options.getAnkiIntegration?.()) {
    return false;
  }

  const config = options.getResolvedConfig();
  const subtitleTimingTracker = options.getSubtitleTimingTracker();
  const mpvClient = options.getMpvClient();
  const runtimeOptionsManager = options.getRuntimeOptionsManager();

  if (
    config.ankiConnect?.enabled !== true ||
    !subtitleTimingTracker ||
    !mpvClient ||
    !runtimeOptionsManager
  ) {
    return false;
  }

  const effectiveAnkiConfig = runtimeOptionsManager.getEffectiveAnkiConnectConfig(
    config.ankiConnect,
  );
  const createAnkiIntegration = options.createAnkiIntegration ?? createDefaultAnkiIntegration;
  const integration = createAnkiIntegration({
    config: effectiveAnkiConfig,
    aiConfig: mergeAiConfig(config.ai, config.ankiConnect?.ai),
    subtitleTimingTracker,
    mpvClient,
    showDesktopNotification: options.showDesktopNotification,
    showOverlayNotification: options.showOverlayNotification,
    createFieldGroupingCallback: options.createFieldGroupingCallback,
    knownWordCacheStatePath: options.getKnownWordCacheStatePath(),
  });
  if (options.shouldStartAnkiIntegration?.() !== false) {
    integration.start();
  }
  options.setAnkiIntegration(integration);
  return true;
}
