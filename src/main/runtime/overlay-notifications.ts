import type { OverlayNotificationPayload } from '../../types';

export type NotificationType = 'osd' | 'system' | 'both' | 'none' | undefined;

function normalizeNotificationPayload(
  payload: string | OverlayNotificationPayload,
): OverlayNotificationPayload {
  if (typeof payload === 'string') {
    const trimmed = payload.trim();
    const spinnerNormalized = trimmed.replace(/\s+[|/\\-]$/, '');
    if (
      trimmed.startsWith('Loading subtitle annotations') ||
      trimmed === 'Overlay loading...' ||
      trimmed === 'Starting...' ||
      trimmed === 'Restarting...'
    ) {
      return {
        kind: 'loading',
        message: spinnerNormalized,
      };
    }
    return {
      kind: 'info',
      message: payload,
    };
  }
  return payload;
}

export function createShowLogicalOsdHandler(deps: {
  showOverlayNotification: (payload: OverlayNotificationPayload) => boolean;
  showMpvOsd: (message: string) => void;
}) {
  return (payload: string | OverlayNotificationPayload): 'overlay' | 'osd' => {
    const normalized = normalizeNotificationPayload(payload);
    if (deps.showOverlayNotification(normalized)) {
      return 'overlay';
    }
    deps.showMpvOsd(normalized.message);
    return 'osd';
  };
}

export function createShowOverlayNotificationHandler(deps: {
  isOverlayRuntimeInitialized: () => boolean;
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => {
    isDestroyed: () => boolean;
    isVisible: () => boolean;
    webContents: {
      isLoading: () => boolean;
      getURL: () => string;
      once: (event: 'did-finish-load', listener: () => void) => void;
      send: (channel: string, payload: OverlayNotificationPayload) => void;
    };
  } | null;
  notificationChannel: string;
}) {
  return (payload: OverlayNotificationPayload): boolean => {
    if (!deps.isOverlayRuntimeInitialized()) {
      return false;
    }
    if (!deps.getVisibleOverlayVisible()) {
      return false;
    }

    const mainWindow = deps.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) {
      return false;
    }

    const send = (): void => {
      mainWindow.webContents.send(deps.notificationChannel, payload);
    };

    if (!mainWindow.webContents.isLoading() && mainWindow.webContents.getURL() !== 'about:blank') {
      send();
      return true;
    }

    mainWindow.webContents.once('did-finish-load', () => {
      if (!mainWindow.isDestroyed() && mainWindow.isVisible()) {
        send();
      }
    });
    return true;
  };
}

export function shouldShowLogicalOsd(type: NotificationType): boolean {
  return type === undefined || type === 'osd' || type === 'both';
}

export function shouldShowDesktopNotification(type: NotificationType): boolean {
  return type === 'system' || type === 'both';
}

export function createConfiguredNotificationHandler(deps: {
  getNotificationType: () => NotificationType;
  showLogicalOsd: (payload: OverlayNotificationPayload) => 'overlay' | 'osd';
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
}) {
  return (title: string, payload: string | OverlayNotificationPayload): void => {
    const normalized = normalizeNotificationPayload(payload);
    const notificationType = deps.getNotificationType();

    if (shouldShowLogicalOsd(notificationType)) {
      deps.showLogicalOsd(normalized);
    }

    if (shouldShowDesktopNotification(notificationType)) {
      deps.showDesktopNotification(title, { body: normalized.message });
    }
  };
}
