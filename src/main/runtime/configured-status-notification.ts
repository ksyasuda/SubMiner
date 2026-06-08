import type { NotificationType, OverlayNotificationPayload } from '../../types/notification';
import { shouldShowDesktop, shouldShowOverlay, shouldShowOsd } from './notification-routing';

export interface ConfiguredStatusNotificationDeps {
  getNotificationType: () => NotificationType | undefined;
  isOverlayReady?: () => boolean;
  showOsd: (message: string) => boolean | void;
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  showDesktopNotification: (title: string, options: { body?: string }) => void;
}

export interface ConfiguredStatusNotificationOptions {
  id?: string;
  title?: string;
  variant?: OverlayNotificationPayload['variant'];
  persistent?: boolean;
  desktop?: boolean;
  delivery?: 'notification' | 'feedback';
}

export function getPlaybackFeedbackNotificationOptions(
  message: string,
): ConfiguredStatusNotificationOptions {
  if (/^Primary subtitle: (hidden|visible|hover)$/.test(message)) {
    return { id: 'primary-subtitle-mode-feedback' };
  }
  if (/^Secondary subtitle: (hidden|visible|hover)$/.test(message)) {
    return { id: 'secondary-subtitle-mode-feedback' };
  }
  return {};
}

export function notifyConfiguredStatus(
  message: string,
  deps: ConfiguredStatusNotificationDeps,
  options: ConfiguredStatusNotificationOptions = {},
): void {
  const type = deps.getNotificationType() ?? 'overlay';
  const delivery = options.delivery ?? 'notification';
  const showOverlay = shouldShowOverlay(type);
  const showOsd = shouldShowOsd(type);
  const desktopEnabled = delivery !== 'feedback' && options.desktop !== false;

  if (type === 'none') {
    return;
  }

  if (delivery === 'feedback' && !showOverlay && !showOsd) {
    return;
  }

  if (showOverlay) {
    const overlayReady = deps.isOverlayReady?.() ?? true;
    if (deps.showOverlayNotification && overlayReady) {
      deps.showOverlayNotification({
        id: options.id,
        title: options.title ?? 'SubMiner',
        body: message,
        variant: options.variant ?? 'info',
        persistent: options.persistent ?? false,
      });
    } else if (desktopEnabled && !shouldShowDesktop(type)) {
      deps.showDesktopNotification(options.title ?? 'SubMiner', { body: message });
    }
  }

  if (showOsd) {
    deps.showOsd(message);
  }

  if (desktopEnabled && shouldShowDesktop(type)) {
    deps.showDesktopNotification(options.title ?? 'SubMiner', { body: message });
  }
}
