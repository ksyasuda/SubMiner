import { ipcRenderer, IpcRendererEvent } from "electron";
import {
  MainToRendererEventChannel,
  RendererToMainInvokeChannel,
  RendererToMainSendChannel,
} from "./contract";

export function invokeFromRenderer<T>(
  channel: RendererToMainInvokeChannel,
  ...args: unknown[]
): Promise<T> {
  return ipcRenderer.invoke(channel, ...args) as Promise<T>;
}

export function sendFromRenderer(
  channel: RendererToMainSendChannel,
  ...args: unknown[]
): void {
  ipcRenderer.send(channel, ...args);
}

export function onMainEvent(
  channel: MainToRendererEventChannel,
  listener: (event: IpcRendererEvent, ...args: unknown[]) => void,
): void {
  ipcRenderer.on(channel, listener);
}
