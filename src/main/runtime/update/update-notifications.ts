import type { UpdateNotificationType } from '../../../types/config';
import type { OverlayNotificationPayload } from '../../../types/notification';
import { i18n } from '../../../i18n/index.js';

export const UPDATE_AVAILABLE_NOTIFICATION_ID = 'subminer-update-available';
export const INSTALL_UPDATE_ACTION_ID = 'install-update';

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

  const message = i18n.t('update.available', { version: options.version });
  if (options.notificationType === 'overlay' || options.notificationType === 'both') {
    deps.showOverlayNotification({
      id: UPDATE_AVAILABLE_NOTIFICATION_ID,
      title: i18n.t('update.updateAvailableTitle'),
      body: message,
      variant: 'info',
      persistent: true,
      actions: [{ id: INSTALL_UPDATE_ACTION_ID, label: i18n.t('update.updateButton') }],
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
    deps.showSystemNotification(i18n.t('update.updateAvailableTitle'), message);
  }
}
