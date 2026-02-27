import { KikuFieldGroupingChoice, KikuFieldGroupingRequestData } from '../../types';
import { createFieldGroupingCallbackRuntime, sendToVisibleOverlayRuntime } from './overlay-bridge';

interface WindowLike {
  isDestroyed: () => boolean;
  webContents: {
    send: (channel: string, payload?: unknown) => void;
  };
}

export interface FieldGroupingOverlayRuntimeOptions<T extends string> {
  getMainWindow: () => WindowLike | null;
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getResolver: () => ((choice: KikuFieldGroupingChoice) => void) | null;
  setResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) => void;
  getRestoreVisibleOverlayOnModalClose: () => Set<T>;
  sendToVisibleOverlay?: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: T },
  ) => boolean;
}

export function createFieldGroupingOverlayRuntime<T extends string>(
  options: FieldGroupingOverlayRuntimeOptions<T>,
): {
  sendToVisibleOverlay: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: T },
  ) => boolean;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
} {
  const sendToVisibleOverlay = (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: T },
  ): boolean => {
    if (options.sendToVisibleOverlay) {
      const wasVisible = options.getVisibleOverlayVisible();
      const sent = options.sendToVisibleOverlay(channel, payload, runtimeOptions);
      if (sent && !wasVisible && !options.getVisibleOverlayVisible()) {
        options.setVisibleOverlayVisible(true);
      }
      return sent;
    }
    return sendToVisibleOverlayRuntime({
      mainWindow: options.getMainWindow() as never,
      visibleOverlayVisible: options.getVisibleOverlayVisible(),
      setVisibleOverlayVisible: options.setVisibleOverlayVisible,
      channel,
      payload,
      restoreOnModalClose: runtimeOptions?.restoreOnModalClose,
      restoreVisibleOverlayOnModalClose: options.getRestoreVisibleOverlayOnModalClose(),
    });
  };

  const createFieldGroupingCallback = (): ((
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>) => {
    return createFieldGroupingCallbackRuntime({
      getVisibleOverlayVisible: options.getVisibleOverlayVisible,
      setVisibleOverlayVisible: options.setVisibleOverlayVisible,
      getResolver: options.getResolver,
      setResolver: options.setResolver,
      sendToVisibleOverlay,
    });
  };

  return {
    sendToVisibleOverlay,
    createFieldGroupingCallback,
  };
}
