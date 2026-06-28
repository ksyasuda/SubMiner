import type { NotificationType, OverlayNotificationPayload } from '../../types/notification';
import { shouldShowDesktop, shouldShowOverlay, shouldShowOsd } from './notification-routing';

export interface ConfiguredStatusNotificationDeps {
  getNotificationType: () => NotificationType | undefined;
  isOverlayReady?: () => boolean;
  showOsd: (message: string) => boolean | void;
  queueOsd?: (message: string, options: ConfiguredStatusNotificationOptions) => void;
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
    if (deps.showOverlayNotification) {
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
    const shown = deps.showOsd(message);
    if (shown === false && delivery !== 'feedback') {
      deps.queueOsd?.(message, options);
    }
  }

  if (desktopEnabled && shouldShowDesktop(type)) {
    deps.showDesktopNotification(options.title ?? 'SubMiner', { body: message });
  }
}
