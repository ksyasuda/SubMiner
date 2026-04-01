import type { OverlayHostedModal } from '../shared/ipc/contracts';
import type { WindowGeometry } from '../types';
import type {
  OverlayUiActionsInput,
  OverlayUiBootstrapInput,
  OverlayUiGeometryInput,
  OverlayUiModalInput,
  OverlayUiMpvSubtitleInput,
  OverlayUiRuntimeInput,
  OverlayUiRuntimeStateInput,
  OverlayUiTrayInput,
  OverlayUiVisibilityActionsInput,
  OverlayUiVisibilityServiceInput,
  OverlayUiWindowState,
  OverlayUiWindowsInput,
} from './overlay-ui-runtime';

type WindowLike = {
  isDestroyed: () => boolean;
};

export interface OverlayUiRuntimeWindowsInput<TWindow extends WindowLike = WindowLike> {
  windowState: OverlayUiWindowState<TWindow>;
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
  visibilityService: OverlayUiVisibilityServiceInput<TWindow>;
  overlayWindows: OverlayUiWindowsInput<TWindow>;
  visibilityActions: OverlayUiVisibilityActionsInput;
}

export interface OverlayUiRuntimeGroupedInput<TWindow extends WindowLike = WindowLike> {
  windows: OverlayUiRuntimeWindowsInput<TWindow>;
  overlayActions: OverlayUiActionsInput;
  tray: OverlayUiTrayInput | null;
  bootstrap: OverlayUiBootstrapInput;
  runtimeState: OverlayUiRuntimeStateInput;
  mpvSubtitle: OverlayUiMpvSubtitleInput;
}

export type OverlayUiRuntimeInputLike<TWindow extends WindowLike = WindowLike> =
  | OverlayUiRuntimeInput<TWindow>
  | OverlayUiRuntimeGroupedInput<TWindow>;

export function normalizeOverlayUiRuntimeInput<TWindow extends WindowLike>(
  input: OverlayUiRuntimeInputLike<TWindow>,
): OverlayUiRuntimeInput<TWindow> {
  if (!('windows' in input)) {
    return input;
  }

  return {
    windowState: input.windows.windowState,
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
    visibilityService: input.windows.visibilityService,
    overlayWindows: input.windows.overlayWindows,
    visibilityActions: input.windows.visibilityActions,
    overlayActions: input.overlayActions,
    tray: input.tray,
    bootstrap: input.bootstrap,
    runtimeState: input.runtimeState,
    mpvSubtitle: input.mpvSubtitle,
  };
}
