import electron from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import { createLogger } from '../../logger';

const { Notification, nativeImage } = electron;
const logger = createLogger('core:notification');

export function resolveDefaultNotificationIconPath(deps: {
  platform: string;
  resourcesPath: string;
  appPath: string;
  dirname: string;
  joinPath: (...parts: string[]) => string;
  fileExists: (path: string) => boolean;
}): string | null {
  const iconNames =
    deps.platform === 'win32'
      ? ['SubMiner.ico', 'SubMiner-square.png', 'SubMiner.png']
      : ['SubMiner-square.png', 'SubMiner.png'];

  const baseDirs = [
    deps.joinPath(deps.resourcesPath, 'assets'),
    deps.joinPath(deps.appPath, 'assets'),
    deps.joinPath(deps.dirname, '..', 'assets'),
    deps.joinPath(deps.dirname, '..', '..', 'assets'),
    deps.joinPath(process.cwd(), 'assets'),
  ];

  for (const baseDir of baseDirs) {
    for (const iconName of iconNames) {
      const candidate = deps.joinPath(baseDir, iconName);
      if (deps.fileExists(candidate)) {
        return candidate;
      }
    }
  }

  return null;
}

function resolveRuntimeDefaultNotificationIconPath(): string | null {
  return resolveDefaultNotificationIconPath({
    platform: process.platform,
    resourcesPath: process.resourcesPath,
    appPath: electron.app?.getAppPath?.() ?? process.cwd(),
    dirname: __dirname,
    joinPath: (...parts) => path.join(...parts),
    fileExists: (candidate) => fs.existsSync(candidate),
  });
}

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

  const icon = options.icon ?? resolveRuntimeDefaultNotificationIconPath() ?? undefined;

  if (icon) {
    const isFilePath =
      typeof icon === 'string' && (icon.startsWith('/') || /^[a-zA-Z]:[\\/]/.test(icon));

    if (isFilePath) {
      if (fs.existsSync(icon)) {
        notificationOptions.icon = icon;
      } else {
        logger.warn('Notification icon file not found', icon);
      }
    } else if (typeof icon === 'string' && icon.startsWith('data:image/')) {
      const base64Data = icon.replace(/^data:image\/\w+;base64,/, '');
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
      notificationOptions.icon = icon;
    }
  }

  const notification = new Notification(notificationOptions);
  notification.show();
}
