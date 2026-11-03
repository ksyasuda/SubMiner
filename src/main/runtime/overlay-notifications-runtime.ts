import type { BrowserWindow } from 'electron';
import type {
  NotificationType,
  OverlayNotificationEventPayload,
  OverlayNotificationPayload,
  ResolvedConfig,
} from '../../types';
import type { AnkiIntegration } from '../../anki-integration';
import type { RuntimeOptionsManager } from '../../runtime-options';
import { AnkiConnectClient } from '../../anki-connect';
import { DEFAULT_CONFIG } from '../../config';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { showDesktopNotification } from '../../core/utils';
import {
  isOverlayWindowContentReady,
  sendMpvCommandRuntime,
  type MpvIpcClient,
} from '../../core/services';
import { createOverlayLoadingOsdController } from './overlay-loading-osd';
import { createMaybeStartOverlayLoadingOsdHandler } from './overlay-loading-osd-start';
import { withConfiguredOverlayNotificationPosition } from './overlay-notification-position';
import { createOverlayNotificationDelivery } from './overlay-notification-delivery';
import {
  getPlaybackFeedbackNotificationOptions,
  getSubsyncStatusNotificationOptions,
  getYoutubeFlowStatusNotificationOptions,
  notifyConfiguredStatus,
  type ConfiguredStatusNotificationOptions,
} from './configured-status-notification';
import { resolveOverlayReadinessNotificationType } from './notification-routing';

export interface OverlayNotificationsRuntimeDeps {
  getResolvedConfig: () => ResolvedConfig;
  getMainOverlayWindow: () => BrowserWindow | null;
  getVisibleOverlayVisible: () => boolean;
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void;
  showMpvOsd: (message: string) => boolean | void;
  getMpvClient: () => MpvIpcClient | null;
  getAnkiIntegration: () => AnkiIntegration | null;
  getRuntimeOptionsManager: () => RuntimeOptionsManager | null;
}

export function createOverlayNotificationsRuntime(deps: OverlayNotificationsRuntimeDeps): {
  isVisibleOverlayContentReady: () => boolean;
  getConfiguredStatusNotificationType: () => NotificationType;
  flushQueuedOverlayNotifications: () => void;
  flushQueuedMpvOsdNotifications: () => void;
  showOverlayNotification: (payload: OverlayNotificationPayload) => void;
  dismissOverlayNotification: (id: string) => void;
  openAnkiCardFromNotification: (noteId: number) => Promise<void>;
  toggleNotificationHistoryPanel: () => void;
  showConfiguredStatusNotification: (
    message: string,
    options?: ConfiguredStatusNotificationOptions,
  ) => void;
  showConfiguredPlaybackFeedback: (
    message: string,
    options?: ConfiguredStatusNotificationOptions,
  ) => void;
  showSubsyncStatusNotification: (message: string) => void;
  showYoutubeFlowStatusNotification: (message: string) => void;
  showOverlayLoadingStatusNotification: () => void;
  dismissOverlayLoadingStatusNotification: () => void;
  reportOverlayContentReadyToPlugin: () => void;
  maybeStartOverlayLoadingOsd: (mediaPath?: string | null) => void;
} {
  function isVisibleOverlayContentReady(): boolean {
    const overlayWindow = deps.getMainOverlayWindow();
    return Boolean(
      deps.getVisibleOverlayVisible() &&
      overlayWindow &&
      isOverlayWindowReadyForNotification(overlayWindow),
    );
  }

  function getConfiguredStatusNotificationType(): NotificationType {
    const configuredType = deps.getResolvedConfig().ankiConnect.behavior.notificationType;
    return resolveOverlayReadinessNotificationType(configuredType, isVisibleOverlayContentReady());
  }

  function isOverlayWindowReadyForNotification(window: BrowserWindow): boolean {
    if (window.isDestroyed() || !isOverlayWindowContentReady(window)) {
      return false;
    }
    if (window.webContents.isLoading()) {
      return false;
    }
    const currentURL = window.webContents.getURL();
    return currentURL !== '' && currentURL !== 'about:blank';
  }

  const overlayNotificationDelivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => isVisibleOverlayContentReady(),
    send: (payload) => {
      deps.broadcastToOverlayWindows(IPC_CHANNELS.event.overlayNotification, payload);
    },
    scheduleFlushRetry: (callback, delayMs) => setTimeout(callback, delayMs),
    clearFlushRetry: (handle) => clearTimeout(handle as ReturnType<typeof setTimeout>),
  });
  let overlayLoadingOsdController: ReturnType<typeof createOverlayLoadingOsdController> | null =
    null;
  const queuedConfiguredOsdNotifications = new Map<
    string,
    { message: string; options: ConfiguredStatusNotificationOptions }
  >();

  function flushQueuedOverlayNotifications(): void {
    overlayNotificationDelivery.flush();
  }

  function queueConfiguredOsdNotification(
    message: string,
    options: ConfiguredStatusNotificationOptions,
  ): void {
    const key = options.id ?? message;
    queuedConfiguredOsdNotifications.set(key, { message, options });
    while (queuedConfiguredOsdNotifications.size > 16) {
      const oldestKey = queuedConfiguredOsdNotifications.keys().next().value;
      if (typeof oldestKey !== 'string') {
        break;
      }
      queuedConfiguredOsdNotifications.delete(oldestKey);
    }
  }

  function flushQueuedMpvOsdNotifications(): void {
    for (const [key, entry] of [...queuedConfiguredOsdNotifications.entries()]) {
      if (deps.showMpvOsd(entry.message) !== false) {
        queuedConfiguredOsdNotifications.delete(key);
      }
    }
  }

  function sendOverlayNotificationEvent(payload: OverlayNotificationEventPayload): void {
    overlayNotificationDelivery.send(payload);
  }

  function showOverlayNotification(payload: OverlayNotificationPayload): void {
    sendOverlayNotificationEvent(
      withConfiguredOverlayNotificationPosition(payload, deps.getResolvedConfig()),
    );
  }

  function dismissOverlayNotification(id: string): void {
    sendOverlayNotificationEvent({ id, dismiss: true });
  }

  async function openAnkiCardFromNotification(noteId: number): Promise<void> {
    const activeIntegrationOpen = deps.getAnkiIntegration()?.openNoteInAnki(noteId);
    if (activeIntegrationOpen) {
      await activeIntegrationOpen;
      return;
    }

    const resolvedConfig = deps.getResolvedConfig();
    const effectiveAnkiConfig =
      deps.getRuntimeOptionsManager()?.getEffectiveAnkiConnectConfig(resolvedConfig.ankiConnect) ??
      resolvedConfig.ankiConnect;
    const fallbackClient = new AnkiConnectClient(
      effectiveAnkiConfig.url || DEFAULT_CONFIG.ankiConnect.url,
    );
    await fallbackClient.openNoteInBrowser(noteId);
  }

  function toggleNotificationHistoryPanel(): void {
    deps.broadcastToOverlayWindows(IPC_CHANNELS.event.notificationHistoryToggle);
  }

  function showConfiguredStatusNotification(
    message: string,
    options: ConfiguredStatusNotificationOptions = {},
  ): void {
    notifyConfiguredStatus(
      message,
      {
        getNotificationType: () => deps.getResolvedConfig().ankiConnect.behavior.notificationType,
        isOverlayReady: () => isVisibleOverlayContentReady(),
        showOsd: (text) => deps.showMpvOsd(text),
        queueOsd: (text, queueOptions) => queueConfiguredOsdNotification(text, queueOptions),
        showOverlayNotification,
        showDesktopNotification: (title, notificationOptions) =>
          showDesktopNotification(title, notificationOptions),
      },
      options,
    );
  }

  function showConfiguredPlaybackFeedback(
    message: string,
    options: ConfiguredStatusNotificationOptions = {},
  ): void {
    showConfiguredStatusNotification(message, {
      ...getPlaybackFeedbackNotificationOptions(message),
      ...options,
      delivery: 'feedback',
    });
  }

  function showSubsyncStatusNotification(message: string): void {
    showConfiguredStatusNotification(message, getSubsyncStatusNotificationOptions(message));
  }

  function showYoutubeFlowStatusNotification(message: string): void {
    showConfiguredStatusNotification(message, getYoutubeFlowStatusNotificationOptions(message));
  }

  function getOverlayLoadingOsdController(): ReturnType<typeof createOverlayLoadingOsdController> {
    if (!overlayLoadingOsdController) {
      overlayLoadingOsdController = createOverlayLoadingOsdController({
        showOsd: (message) => {
          deps.showMpvOsd(message);
        },
        clearOsd: () => {
          sendMpvCommandRuntime(deps.getMpvClient(), ['show-text', '', '1']);
        },
        setInterval: (callback, delayMs) => {
          const timer = setInterval(callback, delayMs);
          timer.unref?.();
          return timer;
        },
        clearInterval: (timer) => {
          clearInterval(timer as ReturnType<typeof setInterval>);
        },
      });
    }
    return overlayLoadingOsdController;
  }

  function showOverlayLoadingStatusNotification(): void {
    getOverlayLoadingOsdController().start();
  }

  function isMainOverlayContentReady(): boolean {
    const overlayWindow = deps.getMainOverlayWindow();
    return Boolean(overlayWindow && isOverlayWindowContentReady(overlayWindow));
  }

  function sendPluginOverlayLoadingMessage(
    message: 'subminer-overlay-loading-ready' | 'subminer-overlay-loading-dismissed',
  ): void {
    sendMpvCommandRuntime(deps.getMpvClient(), ['script-message', message]);
  }

  // Hide paths also dismiss, often before the overlay has content. Only a content-ready window
  // reports `subminer-overlay-loading-ready`: the plugin remembers it and skips the spinner on
  // later playlist files, so other dismissals send the stop-only message, which also makes the
  // plugin forget readiness from a replaced window.
  function dismissOverlayLoadingStatusNotification(): void {
    getOverlayLoadingOsdController().stop();
    sendPluginOverlayLoadingMessage(
      isMainOverlayContentReady()
        ? 'subminer-overlay-loading-ready'
        : 'subminer-overlay-loading-dismissed',
    );
    dismissOverlayNotification('overlay-loading-status');
  }

  // Called when mpv (re)connects: a ready message sent while mpv was away, or before this mpv's
  // plugin existed, was lost. Never sends the dismissal, since nothing is being hidden.
  function reportOverlayContentReadyToPlugin(): void {
    if (isMainOverlayContentReady()) {
      sendPluginOverlayLoadingMessage('subminer-overlay-loading-ready');
    }
  }

  const maybeStartOverlayLoadingOsd = createMaybeStartOverlayLoadingOsdHandler({
    getVisibleOverlayRequested: () => deps.getVisibleOverlayVisible(),
    isOverlayContentReady: () => isVisibleOverlayContentReady(),
    startOverlayLoadingOsd: () => {
      showOverlayLoadingStatusNotification();
    },
  });

  return {
    isVisibleOverlayContentReady,
    getConfiguredStatusNotificationType,
    flushQueuedOverlayNotifications,
    flushQueuedMpvOsdNotifications,
    showOverlayNotification,
    dismissOverlayNotification,
    openAnkiCardFromNotification,
    toggleNotificationHistoryPanel,
    showConfiguredStatusNotification,
    showConfiguredPlaybackFeedback,
    showSubsyncStatusNotification,
    showYoutubeFlowStatusNotification,
    showOverlayLoadingStatusNotification,
    dismissOverlayLoadingStatusNotification,
    reportOverlayContentReadyToPlugin,
    maybeStartOverlayLoadingOsd,
  };
}
