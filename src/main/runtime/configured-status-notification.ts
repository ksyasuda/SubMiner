import type { NotificationType, OverlayNotificationPayload } from '../../types/notification';

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

function shouldShowOverlay(type: NotificationType): boolean {
  return type === 'overlay' || type === 'both';
}

function shouldShowOsd(type: NotificationType): boolean {
  return type === 'osd' || type === 'osd-system';
}

function shouldShowDesktop(type: NotificationType): boolean {
  return type === 'system' || type === 'both' || type === 'osd-system';
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

  if (deps.isOverlayReady?.() === false) {
    deps.showOsd(message);
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
    deps.showOsd(message);
  }

  if (desktopEnabled && shouldShowDesktop(type)) {
    deps.showDesktopNotification(options.title ?? 'SubMiner', { body: message });
  }
}
