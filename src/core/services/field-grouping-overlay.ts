import { KikuFieldGroupingChoice, KikuFieldGroupingRequestData } from '../../types';
import { createFieldGroupingCallbackRuntime, sendToVisibleOverlayRuntime } from './overlay-bridge';

interface WindowLike {
  isDestroyed: () => boolean;
  webContents: {
    send: (channel: string, payload?: unknown) => void;
  };
}

const KIKU_FIELD_GROUPING_MODAL_OPEN_TIMEOUT_MS = 1500;
const KIKU_FIELD_GROUPING_MODAL_RETRY_WARNING =
  'Kiku field grouping modal did not acknowledge modal open on first attempt; retrying dedicated modal window.';

export interface FieldGroupingOverlayRuntimeOptions<T extends string> {
  getMainWindow: () => WindowLike | null;
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getResolver: () => ((choice: KikuFieldGroupingChoice) => void) | null;
  setResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) => void;
  getRestoreVisibleOverlayOnModalClose: () => Set<T>;
  waitForModalOpen?: (modal: T, timeoutMs: number) => Promise<boolean>;
  handleOverlayModalClosed?: (modal: T) => void;
  logWarn?: (message: string) => void;
  sendToVisibleOverlay?: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: T; preferModalWindow?: boolean },
  ) => boolean;
}

export function createFieldGroupingOverlayRuntime<T extends string>(
  options: FieldGroupingOverlayRuntimeOptions<T>,
): {
  sendToVisibleOverlay: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: T; preferModalWindow?: boolean },
  ) => boolean;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
} {
  const sendToVisibleOverlay = (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: T; preferModalWindow?: boolean },
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

  const sendKikuFieldGroupingRequest = async (
    data: KikuFieldGroupingRequestData,
  ): Promise<boolean> => {
    const kikuModal = 'kiku' as T;
    const sendOpen = (): boolean =>
      sendToVisibleOverlay('kiku:field-grouping-request', data, {
        restoreOnModalClose: kikuModal,
        preferModalWindow: true,
      });

    if (!options.waitForModalOpen) {
      return sendOpen();
    }

    if (!sendOpen()) {
      return false;
    }
    if (await options.waitForModalOpen(kikuModal, KIKU_FIELD_GROUPING_MODAL_OPEN_TIMEOUT_MS)) {
      return true;
    }

    options.logWarn?.(KIKU_FIELD_GROUPING_MODAL_RETRY_WARNING);
    if (!sendOpen()) {
      options.handleOverlayModalClosed?.(kikuModal);
      return false;
    }

    const opened = await options.waitForModalOpen(
      kikuModal,
      KIKU_FIELD_GROUPING_MODAL_OPEN_TIMEOUT_MS,
    );
    if (!opened) {
      options.handleOverlayModalClosed?.(kikuModal);
    }
    return opened;
  };

  const dismissModalUi = (): void => {
    const kikuModal = 'kiku' as T;
    // Best-effort: tell the renderer hosting the modal to close its dialog. When the modal
    // lives in the dedicated modal window this is redundant with the teardown below, but it
    // also covers the case where the request was routed into the visible overlay.
    sendToVisibleOverlay('kiku:field-grouping-cancel', undefined, { preferModalWindow: true });
    // Reliable teardown of main-side modal state (restore set, main-overlay passthrough,
    // dedicated modal window). This is what recovers the frozen overlay when a grouping
    // request times out or fails to reach a visible modal.
    options.handleOverlayModalClosed?.(kikuModal);
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
      sendKikuFieldGroupingRequest,
      dismissModalUi,
    });
  };

  return {
    sendToVisibleOverlay,
    createFieldGroupingCallback,
  };
}
