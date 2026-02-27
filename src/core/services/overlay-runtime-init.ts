import { BrowserWindow } from 'electron';
import { AnkiIntegration } from '../../anki-integration';
import { BaseWindowTracker, createWindowTracker } from '../../window-trackers';
import {
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  WindowGeometry,
} from '../../types';

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
  getMpvSocketPath: () => string;
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
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
}): void {
  options.createMainWindow();
  options.registerGlobalShortcuts();

  const windowTracker = createWindowTracker(options.backendOverride, options.getMpvSocketPath());
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
    windowTracker.start();
  }

  const config = options.getResolvedConfig();
  const subtitleTimingTracker = options.getSubtitleTimingTracker();
  const mpvClient = options.getMpvClient();
  const runtimeOptionsManager = options.getRuntimeOptionsManager();

  if (config.ankiConnect && subtitleTimingTracker && mpvClient && runtimeOptionsManager) {
    const effectiveAnkiConfig = runtimeOptionsManager.getEffectiveAnkiConnectConfig(
      config.ankiConnect,
    );
    const integration = new AnkiIntegration(
      effectiveAnkiConfig,
      subtitleTimingTracker as never,
      mpvClient as never,
      (text: string) => {
        if (mpvClient && typeof mpvClient.send === 'function') {
          mpvClient.send({
            command: ['show-text', text, '3000'],
          });
        }
      },
      options.showDesktopNotification,
      options.createFieldGroupingCallback(),
      options.getKnownWordCacheStatePath(),
    );
    integration.start();
    options.setAnkiIntegration(integration);
  }

  options.updateVisibleOverlayVisibility();
}
