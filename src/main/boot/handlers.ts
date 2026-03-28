import { composeOverlayWindowHandlers } from '../runtime/composers/overlay-window-composer';
import {
  composeCliStartupHandlers,
  composeHeadlessStartupHandlers,
  composeIpcRuntimeHandlers,
  composeStartupLifecycleHandlers,
} from '../runtime/composers';

export interface MainBootHandlersParams<TBrowserWindow, TCliArgs, TStartupState, TBootstrapDeps> {
  startupLifecycleDeps: Parameters<typeof composeStartupLifecycleHandlers>[0];
  ipcRuntimeDeps: Parameters<typeof composeIpcRuntimeHandlers>[0];
  cliStartupDeps: Parameters<typeof composeCliStartupHandlers>[0];
  headlessStartupDeps: Parameters<
    typeof composeHeadlessStartupHandlers<TCliArgs, TStartupState, TBootstrapDeps>
  >[0];
  overlayWindowDeps: Parameters<typeof composeOverlayWindowHandlers<TBrowserWindow>>[0];
}

export function createMainBootHandlers<
  TBrowserWindow,
  TCliArgs,
  TStartupState,
  TBootstrapDeps,
>(params: MainBootHandlersParams<TBrowserWindow, TCliArgs, TStartupState, TBootstrapDeps>) {
  return {
    startupLifecycle: composeStartupLifecycleHandlers(params.startupLifecycleDeps),
    ipcRuntime: composeIpcRuntimeHandlers(params.ipcRuntimeDeps),
    cliStartup: composeCliStartupHandlers(params.cliStartupDeps),
    headlessStartup: composeHeadlessStartupHandlers<TCliArgs, TStartupState, TBootstrapDeps>(
      params.headlessStartupDeps,
    ),
    overlayWindow: composeOverlayWindowHandlers<TBrowserWindow>(params.overlayWindowDeps),
  };
}

export const composeBootStartupLifecycleHandlers = composeStartupLifecycleHandlers;
export const composeBootIpcRuntimeHandlers = composeIpcRuntimeHandlers;
export const composeBootCliStartupHandlers = composeCliStartupHandlers;
export const composeBootHeadlessStartupHandlers = composeHeadlessStartupHandlers;
export const composeBootOverlayWindowHandlers = composeOverlayWindowHandlers;
