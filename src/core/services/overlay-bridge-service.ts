import {
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
} from "../../types";
import { createFieldGroupingCallbackService } from "./field-grouping-service";
import { BrowserWindow } from "electron";

export function sendToVisibleOverlayRuntimeService<T extends string>(options: {
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
  if (options.payload === undefined) {
    options.mainWindow.webContents.send(options.channel);
  } else {
    options.mainWindow.webContents.send(options.channel, options.payload);
  }
  return true;
}

export function createFieldGroupingCallbackRuntimeService<T extends string>(
  options: {
    getVisibleOverlayVisible: () => boolean;
    getInvisibleOverlayVisible: () => boolean;
    setVisibleOverlayVisible: (visible: boolean) => void;
    setInvisibleOverlayVisible: (visible: boolean) => void;
    getResolver: () => ((choice: KikuFieldGroupingChoice) => void) | null;
    setResolver: (
      resolver: ((choice: KikuFieldGroupingChoice) => void) | null,
    ) => void;
    sendToVisibleOverlay: (
      channel: string,
      payload?: unknown,
      runtimeOptions?: { restoreOnModalClose?: T },
    ) => boolean;
  },
): (data: KikuFieldGroupingRequestData) => Promise<KikuFieldGroupingChoice> {
  return createFieldGroupingCallbackService({
    getVisibleOverlayVisible: options.getVisibleOverlayVisible,
    getInvisibleOverlayVisible: options.getInvisibleOverlayVisible,
    setVisibleOverlayVisible: options.setVisibleOverlayVisible,
    setInvisibleOverlayVisible: options.setInvisibleOverlayVisible,
    getResolver: options.getResolver,
    setResolver: options.setResolver,
    sendRequestToVisibleOverlay: (data) =>
      options.sendToVisibleOverlay("kiku:field-grouping-request", data),
  });
}
