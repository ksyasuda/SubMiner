import { ipcMain, IpcMainEvent } from "electron";
import {
  RendererToMainInvokeChannel,
  RendererToMainSendChannel,
} from "./contract";

export function onRendererSend(
  channel: RendererToMainSendChannel,
  listener: (event: IpcMainEvent, ...args: any[]) => void,
): void {
  ipcMain.on(channel, listener);
}

export function handleRendererInvoke(
  channel: RendererToMainInvokeChannel,
  handler: (event: Electron.IpcMainInvokeEvent, ...args: any[]) => unknown,
): void {
  ipcMain.handle(channel, handler);
}
