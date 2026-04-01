import type { OverlayHostedModal } from '../shared/ipc/contracts';
import type {
  OverlayUiActionsInput,
  OverlayUiBootstrapInput,
  OverlayUiGeometryInput,
  OverlayUiModalInput,
  OverlayUiMpvSubtitleInput,
  OverlayUiRuntimeStateInput,
  OverlayUiTrayInput,
  OverlayUiVisibilityActionsInput,
  OverlayUiVisibilityServiceInput,
  OverlayUiWindowState,
  OverlayUiWindowsInput,
} from './overlay-ui-runtime';
import type { OverlayUiRuntimeGroupedInput } from './overlay-ui-runtime-input';

type WindowLike = {
  isDestroyed: () => boolean;
};

export interface OverlayUiBootstrapRuntimeWindowsInput<TWindow extends WindowLike = WindowLike> {
  state: OverlayUiWindowState<TWindow>;
  geometry: OverlayUiGeometryInput;
  modal: OverlayUiModalInput;
  modalRuntime: {
    handleOverlayModalClosed: (modal: OverlayHostedModal) => void;
    notifyOverlayModalOpened: (modal: OverlayHostedModal) => void;
    waitForModalOpen: (modal: OverlayHostedModal, timeoutMs: number) => Promise<boolean>;
    getRestoreVisibleOverlayOnModalClose: () => Set<OverlayHostedModal>;
    openRuntimeOptionsPalette: () => void;
    sendToActiveOverlayWindow: (
      channel: string,
      payload?: unknown,
      runtimeOptions?: {
        restoreOnModalClose?: OverlayHostedModal;
        preferModalWindow?: boolean;
      },
    ) => boolean;
  };
  visibility: {
    service: OverlayUiVisibilityServiceInput<TWindow>;
    overlayWindows: OverlayUiWindowsInput<TWindow>;
    actions: OverlayUiVisibilityActionsInput;
  };
}

export interface OverlayUiBootstrapRuntimeInput<TWindow extends WindowLike = WindowLike> {
  windows: OverlayUiBootstrapRuntimeWindowsInput<TWindow>;
  overlayActions: OverlayUiActionsInput;
  tray: OverlayUiTrayInput | null;
  bootstrap: OverlayUiBootstrapInput;
  runtimeState: OverlayUiRuntimeStateInput;
  mpvSubtitle: OverlayUiMpvSubtitleInput;
}

export function createOverlayUiBootstrapRuntimeInput<TWindow extends WindowLike>(
  input: OverlayUiBootstrapRuntimeInput<TWindow>,
): OverlayUiRuntimeGroupedInput<TWindow> {
  return {
    windows: {
      windowState: input.windows.state,
      geometry: input.windows.geometry,
      modal: input.windows.modal,
      modalRuntime: {
        handleOverlayModalClosed: (modal) =>
          input.windows.modalRuntime.handleOverlayModalClosed(modal),
        notifyOverlayModalOpened: (modal) =>
          input.windows.modalRuntime.notifyOverlayModalOpened(modal),
        waitForModalOpen: (modal, timeoutMs) =>
          input.windows.modalRuntime.waitForModalOpen(modal, timeoutMs),
        getRestoreVisibleOverlayOnModalClose: () =>
          input.windows.modalRuntime.getRestoreVisibleOverlayOnModalClose(),
        openRuntimeOptionsPalette: () => input.windows.modalRuntime.openRuntimeOptionsPalette(),
        sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
          input.windows.modalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
      },
      visibilityService: input.windows.visibility.service,
      overlayWindows: input.windows.visibility.overlayWindows,
      visibilityActions: input.windows.visibility.actions,
    },
    overlayActions: input.overlayActions,
    tray: input.tray,
    bootstrap: input.bootstrap,
    runtimeState: input.runtimeState,
    mpvSubtitle: input.mpvSubtitle,
  };
}
