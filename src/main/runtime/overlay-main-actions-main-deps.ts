import type {
  createAppendClipboardVideoToQueueHandler,
  createHandleOverlayModalClosedHandler,
  createSetOverlayVisibleHandler,
  createToggleOverlayHandler,
} from './overlay-main-actions';

type SetOverlayVisibleMainDeps = Parameters<typeof createSetOverlayVisibleHandler>[0];
type ToggleOverlayMainDeps = Parameters<typeof createToggleOverlayHandler>[0];
type HandleOverlayModalClosedMainDeps = Parameters<typeof createHandleOverlayModalClosedHandler>[0];
type AppendClipboardVideoToQueueMainDeps = Parameters<typeof createAppendClipboardVideoToQueueHandler>[0];

export function createBuildSetOverlayVisibleMainDepsHandler(deps: SetOverlayVisibleMainDeps) {
  return (): SetOverlayVisibleMainDeps => ({
    setVisibleOverlayVisible: (visible: boolean) => deps.setVisibleOverlayVisible(visible),
  });
}

export function createBuildToggleOverlayMainDepsHandler(deps: ToggleOverlayMainDeps) {
  return (): ToggleOverlayMainDeps => ({
    toggleVisibleOverlay: () => deps.toggleVisibleOverlay(),
  });
}

export function createBuildHandleOverlayModalClosedMainDepsHandler(
  deps: HandleOverlayModalClosedMainDeps,
) {
  return (): HandleOverlayModalClosedMainDeps => ({
    handleOverlayModalClosedRuntime: (modal) => deps.handleOverlayModalClosedRuntime(modal),
  });
}

export function createBuildAppendClipboardVideoToQueueMainDepsHandler(
  deps: AppendClipboardVideoToQueueMainDeps,
) {
  return (): AppendClipboardVideoToQueueMainDeps => ({
    appendClipboardVideoToQueueRuntime: (options) => deps.appendClipboardVideoToQueueRuntime(options),
    getMpvClient: () => deps.getMpvClient(),
    readClipboardText: () => deps.readClipboardText(),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    sendMpvCommand: (command: (string | number)[]) => deps.sendMpvCommand(command),
  });
}
