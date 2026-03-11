import electron from 'electron';
import type { BrowserWindow, Extension, Session } from 'electron';
import { createLogger } from '../../logger';

const { BrowserWindow: ElectronBrowserWindow, session } = electron;
const logger = createLogger('main:yomitan-settings');

export interface OpenYomitanSettingsWindowOptions {
  yomitanExt: Extension | null;
  getExistingWindow: () => BrowserWindow | null;
  setWindow: (window: BrowserWindow | null) => void;
  yomitanSession?: Session | null;
  onWindowClosed?: () => void;
}

export function openYomitanSettingsWindow(options: OpenYomitanSettingsWindowOptions): void {
  logger.info('openYomitanSettings called');

  if (!options.yomitanExt) {
    logger.error('Yomitan extension not loaded - yomitanExt is:', options.yomitanExt);
    logger.error('This may be due to Manifest V3 service worker issues with Electron');
    return;
  }

  const existingWindow = options.getExistingWindow();
  if (existingWindow && !existingWindow.isDestroyed()) {
    logger.info('Settings window already exists, focusing');
    existingWindow.focus();
    return;
  }

  logger.info('Creating new settings window for extension:', options.yomitanExt.id);

  const settingsWindow = new ElectronBrowserWindow({
    width: 1200,
    height: 800,
    show: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      session: options.yomitanSession ?? session.defaultSession,
    },
  });
  options.setWindow(settingsWindow);

  const settingsUrl = `chrome-extension://${options.yomitanExt.id}/settings.html`;
  logger.info('Loading settings URL:', settingsUrl);

  let loadAttempts = 0;
  const maxAttempts = 3;

  const attemptLoad = (): void => {
    settingsWindow
      .loadURL(settingsUrl)
      .then(() => {
        logger.info('Settings URL loaded successfully');
      })
      .catch((err: Error) => {
        logger.error('Failed to load settings URL:', err);
        loadAttempts++;
        if (loadAttempts < maxAttempts && !settingsWindow.isDestroyed()) {
          logger.info(`Retrying in 500ms (attempt ${loadAttempts + 1}/${maxAttempts})`);
          setTimeout(attemptLoad, 500);
        }
      });
  };

  attemptLoad();

  settingsWindow.webContents.on('did-fail-load', (_event, errorCode, errorDescription) => {
    logger.error('Settings page failed to load:', errorCode, errorDescription);
  });

  settingsWindow.webContents.on('did-finish-load', () => {
    logger.info('Settings page loaded successfully');
  });

  setTimeout(() => {
    if (!settingsWindow.isDestroyed()) {
      const [width = 0, height = 0] = settingsWindow.getSize();
      settingsWindow.setSize(width, height);
      settingsWindow.webContents.invalidate();
      settingsWindow.show();
    }
  }, 500);

  settingsWindow.on('closed', () => {
    options.onWindowClosed?.();
    options.setWindow(null);
  });
}
