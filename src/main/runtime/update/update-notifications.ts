import type { UpdateNotificationType } from '../../../types/config';
import type { OverlayNotificationPayload } from '../../../types/notification';

export interface UpdateNotificationDeps {
  showSystemNotification: (title: string, body: string) => void;
  showOverlayNotification: (payload: OverlayNotificationPayload) => void;
  showOsdNotification: (message: string) => void | Promise<void>;
  log: (message: string) => void;
}

export async function notifyUpdateAvailable(
  options: { notificationType: UpdateNotificationType; version: string },
  deps: UpdateNotificationDeps,
): Promise<void> {
  if (options.notificationType === 'none') return;

  const message = `SubMiner v${options.version} is available`;
  if (options.notificationType === 'overlay' || options.notificationType === 'both') {
    deps.showOverlayNotification({
      title: 'SubMiner update available',
      body: message,
      variant: 'info',
    });
  }
  if (options.notificationType === 'osd' || options.notificationType === 'osd-system') {
    try {
      await deps.showOsdNotification(message);
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      deps.log(`Update OSD notification failed: ${reason}`);
    }
  }
  if (
    options.notificationType === 'system' ||
    options.notificationType === 'both' ||
    options.notificationType === 'osd-system'
  ) {
    deps.showSystemNotification('SubMiner update available', message);
  }
}
