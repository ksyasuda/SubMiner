import type { UpdateNotificationType } from '../../../types/config';

export interface UpdateNotificationDeps {
  showSystemNotification: (title: string, body: string) => void;
  showOsdNotification: (message: string) => void | Promise<void>;
  log: (message: string) => void;
}

export async function notifyUpdateAvailable(
  options: { notificationType: UpdateNotificationType; version: string },
  deps: UpdateNotificationDeps,
): Promise<void> {
  if (options.notificationType === 'none') return;

  const message = `SubMiner v${options.version} is available`;
  if (options.notificationType === 'system' || options.notificationType === 'both') {
    deps.showSystemNotification('SubMiner update available', message);
  }
  if (options.notificationType === 'osd' || options.notificationType === 'both') {
    try {
      await deps.showOsdNotification(message);
    } catch (error) {
      deps.log(`Update OSD notification failed: ${(error as Error).message}`);
    }
  }
}
