import path from 'node:path';
import type { BrowserWindow, BrowserWindowConstructorOptions } from 'electron';
import { AnkiConnectClient } from '../../anki-connect';
import type { ConfigService } from '../../config/service';
import {
  sendMpvCommandRuntime,
  toSubtitleClock,
  type createOverlayManager,
} from '../../core/services';
import { createPluginHost } from '../../core/services/plugins/plugin-host';
import { buildPluginSubtitleContext } from '../../core/services/plugins/subtitle-context';
import type { PluginSubtitleLine } from '../../core/services/plugins/types';
import type { Logger } from '../../logger';
import type { SessionActionDispatchRequest } from '../../types/runtime';
import type { OverlayModalRuntime } from '../overlay-runtime';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createPluginPanelRuntime, openPluginPanelModal } from '../runtime/plugin-panel-runtime';
import { createPluginWindowRuntime } from '../runtime/plugin-window-runtime';
import { resolvePrimarySubtitle } from '../runtime/primary-subtitle-text';
import type { AppState } from '../state';
import type { OverlayModalsAssembly } from './overlay-modals';

type OverlayManager = ReturnType<typeof createOverlayManager>;

/**
 * What the plugin host needs from the rest of `main.ts`. `appState`, `configService` and
 * `overlayManager` are also handed to plugins whole as the `unstable` API.
 */
export interface PluginsAssemblyDeps {
  appState: AppState;
  configService: ConfigService;
  overlayManager: OverlayManager;
  logger: Pick<Logger, 'error'>;
  createLogger: (scope: string) => Logger;
  configDir: string;
  userDataPath: string;
  overlayModalRuntime: Pick<OverlayModalRuntime, 'handleOverlayModalClosed'>;
  overlayModals: Pick<OverlayModalsAssembly, 'createOverlayHostedModalOpenDeps'>;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'showConfiguredStatusNotification'
  >;
  createBrowserWindow: (options: BrowserWindowConstructorOptions) => BrowserWindow;
  /** Keeps a plugin window above the overlay, like the settings windows. */
  promoteAboveOverlay: (window: BrowserWindow) => void;
  /** Starts mpv for a plugin that needs playback (the Jellyfin launch path). */
  ensureMpvRunning: () => Promise<boolean>;
  dispatchSessionAction: (request: SessionActionDispatchRequest) => Promise<void>;
}

export type PluginsAssembly = ReturnType<typeof createPluginsAssembly>;

/**
 * Owns user plugins: the plugin host, the overlay-hosted plugin panel and plugin windows.
 * `main.ts` builds it once; startup loads the plugins and quit cleanup unloads them.
 */
export function createPluginsAssembly(deps: PluginsAssemblyDeps) {
  const { appState, configService, overlayManager, logger } = deps;

  const panel = createPluginPanelRuntime({
    openModal: (getSnapshot, signal) =>
      openPluginPanelModal(
        deps.overlayModals.createOverlayHostedModalOpenDeps(),
        getSnapshot,
        signal,
      ),
    closeModal: () => deps.overlayModalRuntime.handleOverlayModalClosed('plugin-panel'),
    // The panel may live in either window; broadcastToOverlayWindows only reaches the main one.
    broadcast: (channel, payload) => {
      for (const window of [overlayManager.getMainWindow(), overlayManager.getModalWindow()]) {
        if (window && !window.isDestroyed()) window.webContents.send(channel, payload);
      }
    },
    onOpenFailed: () =>
      deps.getOverlayNotifications().showConfiguredStatusNotification('Plugin panel unavailable.', {
        variant: 'warning',
      }),
  });

  /** The line on screen, with cue timings on the subtitle file's timeline (as mining sees it). */
  function getCurrentSubtitle(): PluginSubtitleLine | null {
    const text = appState.currentSubText.trim();
    if (!text) return null;
    const resolved = resolvePrimarySubtitle({
      liveText: appState.mpvClient?.currentSubText ?? '',
      currentTimeSec: toSubtitleClock(
        Number(appState.mpvClient?.currentTimePos),
        appState.mpvClient?.currentSubDelay ?? 0,
      ),
      cues: appState.activeParsedSubtitleCues,
    });
    const resolvedText = resolved?.text.replace(/\n{2,}/g, '\n').trim();
    if (
      resolved &&
      resolvedText &&
      Number.isFinite(resolved.startTime) &&
      Number.isFinite(resolved.endTime) &&
      resolved.endTime > resolved.startTime
    ) {
      return { text: resolvedText, startTime: resolved.startTime, endTime: resolved.endTime };
    }
    const client = appState.mpvClient;
    // The client reports a cleared sub-start as zero, so only trust timing for a live line.
    const hasTiming = client !== null && client.currentSubEnd > client.currentSubStart;
    return {
      text,
      startTime: hasTiming ? client.currentSubStart : null,
      endTime: hasTiming ? client.currentSubEnd : null,
    };
  }

  const windows = createPluginWindowRuntime({
    createBrowserWindow: (options) => deps.createBrowserWindow(options),
    promoteAboveOverlay: (window) => deps.promoteAboveOverlay(window),
    logError: (message) => logger.error(message),
  });

  const host = createPluginHost({
    pluginsDir: path.join(deps.configDir, 'plugins'),
    pluginDataDir: path.join(deps.userDataPath, 'plugin-data'),
    getPluginConfig: () => configService.getConfig().plugins,
    createLogger: deps.createLogger,
    isMpvConnected: () => appState.mpvClient?.connected === true,
    sendMpvCommand: (command) => sendMpvCommandRuntime(appState.mpvClient, command),
    requestMpvProperty: async (name) => {
      if (!appState.mpvClient) throw new Error('mpv is not connected.');
      return await appState.mpvClient.requestProperty(name);
    },
    requestMpvCommand: async (command) => {
      if (!appState.mpvClient) throw new Error('mpv is not connected.');
      return await appState.mpvClient.requestCommand(command);
    },
    ensureMpvRunning: () => deps.ensureMpvRunning(),
    invokeAnkiConnect: (action, params, options) => {
      const ankiConfig = configService.getConfig().ankiConnect;
      const url = (
        appState.runtimeOptionsManager?.getEffectiveAnkiConnectConfig(ankiConfig) ?? ankiConfig
      ).url;
      if (!url) return Promise.reject(new Error('ankiConnect.url is not configured.'));
      // No retries: a retried `sync` or `addNote` could run twice.
      return new AnkiConnectClient(url).invoke(action, params, {
        timeout: options.timeoutMs,
        maxRetries: 0,
      });
    },
    dispatchSessionAction: (request) => deps.dispatchSessionAction(request),
    getCurrentSubtitle: () => getCurrentSubtitle(),
    getSubtitleContext: (options) =>
      buildPluginSubtitleContext({
        cues: appState.activeParsedSubtitleCues,
        current: getCurrentSubtitle(),
        recentLines: (appState.subtitleTimingTracker?.getRecentEntries(40) ?? []).map((entry) => ({
          text: entry.displayText,
          startTime: entry.startTime,
          endTime: entry.endTime,
        })),
        ...options,
      }),
    notify: (message, variant) =>
      deps.getOverlayNotifications().showConfiguredStatusNotification(message, { variant }),
    openPanel: (content) => panel.open(content),
    openWindow: (request) => windows.open(request),
    unstable: { appState, configService, overlayManager },
  });

  return { host, panel, getCurrentSubtitle };
}
