import { KikuFieldGroupingChoice, KikuFieldGroupingRequestData } from '../../types';
import { createFieldGroupingCallback } from './field-grouping';
import type { BrowserWindow } from 'electron';

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
  const webContents = options.mainWindow.webContents as Electron.WebContents & {
    isLoading?: () => boolean;
    getURL?: () => string;
    once?: (event: 'did-finish-load', listener: () => void) => void;
  };
  if (!options.visibleOverlayVisible) {
    options.setVisibleOverlayVisible(true);
  }
  if (!wasVisible && options.restoreOnModalClose) {
    options.restoreVisibleOverlayOnModalClose.add(options.restoreOnModalClose);
  }
  const sendNow = (): void => {
    if (options.payload === undefined) {
      webContents.send(options.channel);
    } else {
      webContents.send(options.channel, options.payload);
    }
  };

  const isLoading = typeof webContents.isLoading === 'function' ? webContents.isLoading() : false;
  const currentURL = typeof webContents.getURL === 'function' ? webContents.getURL() : '';
  const isReady = !isLoading && currentURL !== '' && currentURL !== 'about:blank';

  if (!isReady) {
    if (typeof webContents.once !== 'function') {
      sendNow();
      return true;
    }
    webContents.once('did-finish-load', () => {
      if (!options.mainWindow || options.mainWindow.isDestroyed()) return;
      if (typeof webContents.isLoading !== 'function' || !webContents.isLoading()) {
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
    runtimeOptions?: { restoreOnModalClose?: T; preferModalWindow?: boolean },
  ) => boolean;
  sendKikuFieldGroupingRequest?: (data: KikuFieldGroupingRequestData) => Promise<boolean>;
  dismissModalUi?: () => void;
}): (data: KikuFieldGroupingRequestData) => Promise<KikuFieldGroupingChoice> {
  return createFieldGroupingCallback({
    getVisibleOverlayVisible: options.getVisibleOverlayVisible,
    setVisibleOverlayVisible: options.setVisibleOverlayVisible,
    getResolver: options.getResolver,
    setResolver: options.setResolver,
    dismissModalUi: options.dismissModalUi,
    sendRequestToVisibleOverlay: (data) =>
      options.sendKikuFieldGroupingRequest
        ? options.sendKikuFieldGroupingRequest(data)
        : options.sendToVisibleOverlay('kiku:field-grouping-request', data, {
            restoreOnModalClose: 'kiku' as T,
          }),
  });
}
