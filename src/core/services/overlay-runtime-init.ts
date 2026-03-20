import type { BrowserWindow } from 'electron';
import { BaseWindowTracker, createWindowTracker } from '../../window-trackers';
import { mergeAiConfig } from '../../ai/config';
import {
  AiConfig,
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
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
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  knownWordCacheStatePath: string;
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
  );
}

export function initializeOverlayRuntime(options: {
  backendOverride: string | null;
  createMainWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  isVisibleOverlayVisible: () => boolean;
  updateVisibleOverlayVisibility: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  setWindowTracker: (tracker: BaseWindowTracker | null) => void;
  createWindowTracker?: (
    override?: string | null,
    targetMpvSocketPath?: string | null,
  ) => BaseWindowTracker | null;
  getMpvSocketPath: () => string;
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig; ai?: AiConfig };
  getSubtitleTimingTracker: () => unknown | null;
  getMpvClient: () => {
    send?: (payload: { command: string[] }) => void;
  } | null;
  getRuntimeOptionsManager: () => {
    getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
  } | null;
  setAnkiIntegration: (integration: unknown | null) => void;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  getKnownWordCacheStatePath: () => string;
  shouldStartAnkiIntegration?: () => boolean;
  createAnkiIntegration?: (args: CreateAnkiIntegrationArgs) => AnkiIntegrationLike;
}): void {
  options.createMainWindow();
  options.registerGlobalShortcuts();

  const createWindowTrackerHandler = options.createWindowTracker ?? createWindowTracker;
  const windowTracker = createWindowTrackerHandler(
    options.backendOverride,
    options.getMpvSocketPath(),
  );
  options.setWindowTracker(windowTracker);
  if (windowTracker) {
    windowTracker.onGeometryChange = (geometry: WindowGeometry) => {
      options.updateVisibleOverlayBounds(geometry);
    };
    windowTracker.onWindowFound = (geometry: WindowGeometry) => {
      options.updateVisibleOverlayBounds(geometry);
      if (options.isVisibleOverlayVisible()) {
        options.updateVisibleOverlayVisibility();
      }
    };
    windowTracker.onWindowLost = () => {
      for (const window of options.getOverlayWindows()) {
        window.hide();
      }
      options.syncOverlayShortcuts();
    };
    windowTracker.onWindowFocusChange = () => {
      if (options.isVisibleOverlayVisible()) {
        options.updateVisibleOverlayVisibility();
      }
      options.syncOverlayShortcuts();
    };
    windowTracker.start();
  }

  const config = options.getResolvedConfig();
  const subtitleTimingTracker = options.getSubtitleTimingTracker();
  const mpvClient = options.getMpvClient();
  const runtimeOptionsManager = options.getRuntimeOptionsManager();

  if (
    config.ankiConnect?.enabled === true &&
    subtitleTimingTracker &&
    mpvClient &&
    runtimeOptionsManager
  ) {
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
      createFieldGroupingCallback: options.createFieldGroupingCallback,
      knownWordCacheStatePath: options.getKnownWordCacheStatePath(),
    });
    if (options.shouldStartAnkiIntegration?.() !== false) {
      integration.start();
    }
    options.setAnkiIntegration(integration);
  }

  options.updateVisibleOverlayVisibility();
}
