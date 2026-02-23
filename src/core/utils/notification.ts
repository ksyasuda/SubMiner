import { Notification, nativeImage } from 'electron';
import * as fs from 'fs';
import { createLogger } from '../../logger';

const logger = createLogger('core:notification');

export function showDesktopNotification(
  title: string,
  options: { body?: string; icon?: string },
): void {
  const notificationOptions: {
    title: string;
    body?: string;
    icon?: Electron.NativeImage | string;
  } = { title };

  if (options.body) {
    notificationOptions.body = options.body;
  }

  if (options.icon) {
    const isFilePath =
      typeof options.icon === 'string' &&
      (options.icon.startsWith('/') || /^[a-zA-Z]:[\\/]/.test(options.icon));

    if (isFilePath) {
      if (fs.existsSync(options.icon)) {
        notificationOptions.icon = options.icon;
      } else {
        logger.warn('Notification icon file not found', options.icon);
      }
    } else if (typeof options.icon === 'string' && options.icon.startsWith('data:image/')) {
      const base64Data = options.icon.replace(/^data:image\/\w+;base64,/, '');
      try {
        const image = nativeImage.createFromBuffer(Buffer.from(base64Data, 'base64'));
        if (image.isEmpty()) {
          logger.warn(
            'Notification icon created from base64 is empty - image format may not be supported by Electron',
          );
        } else {
          notificationOptions.icon = image;
        }
      } catch (err) {
        logger.error('Failed to create notification icon from base64', err);
      }
    } else {
      notificationOptions.icon = options.icon;
    }
  }

  const notification = new Notification(notificationOptions);
  notification.show();
}
