import type { OverlayHostedModal } from '../overlay-runtime';
import type { AppendClipboardVideoToQueueRuntimeDeps } from './clipboard-queue';

export function createSetOverlayVisibleHandler(deps: {
  setVisibleOverlayVisible: (visible: boolean) => void;
}) {
  return (visible: boolean): void => {
    deps.setVisibleOverlayVisible(visible);
  };
}

export function createToggleOverlayHandler(deps: {
  toggleVisibleOverlay: () => void;
}) {
  return (): void => {
    deps.toggleVisibleOverlay();
  };
}

export function createHandleOverlayModalClosedHandler(deps: {
  handleOverlayModalClosedRuntime: (modal: OverlayHostedModal) => void;
}) {
  return (modal: OverlayHostedModal): void => {
    deps.handleOverlayModalClosedRuntime(modal);
  };
}

export function createAppendClipboardVideoToQueueHandler(deps: {
  appendClipboardVideoToQueueRuntime: (
    options: AppendClipboardVideoToQueueRuntimeDeps,
  ) => { ok: boolean; message: string };
  getMpvClient: () => AppendClipboardVideoToQueueRuntimeDeps['getMpvClient'] extends () => infer T
    ? T
    : never;
  readClipboardText: () => string;
  showMpvOsd: (text: string) => void;
  sendMpvCommand: (command: (string | number)[]) => void;
}) {
  return (): { ok: boolean; message: string } => {
    return deps.appendClipboardVideoToQueueRuntime({
      getMpvClient: () => deps.getMpvClient(),
      readClipboardText: deps.readClipboardText,
      showMpvOsd: deps.showMpvOsd,
      sendMpvCommand: deps.sendMpvCommand,
    });
  };
}
