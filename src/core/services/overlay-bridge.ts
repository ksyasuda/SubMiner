import { KikuFieldGroupingChoice, KikuFieldGroupingRequestData } from '../../types';
import { createFieldGroupingCallback } from './field-grouping';
import { BrowserWindow } from 'electron';

export function sendToVisibleOverlayRuntime<T extends string>(options: {
  mainWindow: BrowserWindow | null;
  visibleOverlayVisible: boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  channel: string;
  payload?: unknown;
  restoreOnModalClose?: T;
  restoreVisibleOverlayOnModalClose: Set<T>;
}): boolean {
  if (!options.mainWindow || options.mainWindow.isDestroyed()) return false;
  const wasVisible = options.visibleOverlayVisible;
  if (!options.visibleOverlayVisible) {
    options.setVisibleOverlayVisible(true);
  }
  if (!wasVisible && options.restoreOnModalClose) {
    options.restoreVisibleOverlayOnModalClose.add(options.restoreOnModalClose);
  }
  const sendNow = (): void => {
    if (options.payload === undefined) {
      options.mainWindow!.webContents.send(options.channel);
    } else {
      options.mainWindow!.webContents.send(options.channel, options.payload);
    }
  };

  const getURL = options.mainWindow.webContents.getURL;
  const currentURL =
    typeof getURL === 'function' ? getURL.call(options.mainWindow.webContents) : 'ready';
  const isReady =
    !options.mainWindow.webContents.isLoading() &&
    currentURL !== '' &&
    currentURL !== 'about:blank';

  if (!isReady) {
    options.mainWindow.webContents.once('did-finish-load', () => {
      if (!options.mainWindow || options.mainWindow.isDestroyed()) return;
      if (!options.mainWindow.webContents.isLoading()) {
        sendNow();
      }
    });
    return true;
  }

  sendNow();
  return true;
}

export function createFieldGroupingCallbackRuntime<T extends string>(options: {
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getResolver: () => ((choice: KikuFieldGroupingChoice) => void) | null;
  setResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) => void;
  sendToVisibleOverlay: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: T },
  ) => boolean;
}): (data: KikuFieldGroupingRequestData) => Promise<KikuFieldGroupingChoice> {
  return createFieldGroupingCallback({
    getVisibleOverlayVisible: options.getVisibleOverlayVisible,
    setVisibleOverlayVisible: options.setVisibleOverlayVisible,
    getResolver: options.getResolver,
    setResolver: options.setResolver,
    sendRequestToVisibleOverlay: (data) =>
      options.sendToVisibleOverlay('kiku:field-grouping-request', data, {
        restoreOnModalClose: 'kiku' as T,
      }),
  });
}
