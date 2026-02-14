import { BrowserWindow } from "electron";
import { AnkiIntegration } from "../../anki-integration";
import { BaseWindowTracker, createWindowTracker } from "../../window-trackers";
import {
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  WindowGeometry,
} from "../../types";

export function initializeOverlayRuntimeService(options: {
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
}): {
  invisibleOverlayVisible: boolean;
} {
  options.createMainWindow();
  options.createInvisibleWindow();
  const invisibleOverlayVisible = options.getInitialInvisibleOverlayVisibility();
  options.registerGlobalShortcuts();

  const windowTracker = createWindowTracker(
    options.backendOverride,
    options.getMpvSocketPath(),
  );
  options.setWindowTracker(windowTracker);
  if (windowTracker) {
    windowTracker.onGeometryChange = (geometry: WindowGeometry) => {
      options.updateVisibleOverlayBounds(geometry);
      options.updateInvisibleOverlayBounds(geometry);
    };
    windowTracker.onWindowFound = (geometry: WindowGeometry) => {
      options.updateVisibleOverlayBounds(geometry);
      options.updateInvisibleOverlayBounds(geometry);
      if (options.isVisibleOverlayVisible()) {
        options.updateVisibleOverlayVisibility();
      }
      if (options.isInvisibleOverlayVisible()) {
        options.updateInvisibleOverlayVisibility();
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

  if (
    config.ankiConnect &&
    subtitleTimingTracker &&
    mpvClient &&
    runtimeOptionsManager
  ) {
    const effectiveAnkiConfig =
      runtimeOptionsManager.getEffectiveAnkiConnectConfig(config.ankiConnect);
    const integration = new AnkiIntegration(
      effectiveAnkiConfig,
      subtitleTimingTracker as never,
      mpvClient as never,
      (text: string) => {
        if (mpvClient && typeof mpvClient.send === "function") {
          mpvClient.send({
            command: ["show-text", text, "3000"],
          });
        }
      },
      options.showDesktopNotification,
      options.createFieldGroupingCallback(),
    );
    integration.start();
    options.setAnkiIntegration(integration);
  }

  options.updateVisibleOverlayVisibility();
  options.updateInvisibleOverlayVisibility();

  return { invisibleOverlayVisible };
}
