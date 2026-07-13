import { contextBridge, ipcRenderer } from 'electron';
import { IPC_CHANNELS } from './shared/ipc/contracts';

const statsAPI = {
  confirmNativeDialog: (message: string): boolean => {
    return ipcRenderer.sendSync(IPC_CHANNELS.command.statsNativeConfirmDialog, message) === true;
  },

  beginNativeDialog: (): void => {
    ipcRenderer.sendSync(IPC_CHANNELS.command.statsNativeDialogOpened);
  },

  endNativeDialog: (): void => {
    ipcRenderer.send(IPC_CHANNELS.command.statsNativeDialogClosed);
  },
};

contextBridge.exposeInMainWorld('electronAPI', { stats: statsAPI });
