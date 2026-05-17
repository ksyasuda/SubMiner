import { contextBridge, ipcRenderer } from 'electron';
import type {
  ConfigSettingsAPI,
  ConfigSettingsPatch,
  ConfigSettingsSaveResult,
  ConfigSettingsSnapshot,
} from './types/settings';

const SETTINGS_IPC_CHANNELS = {
  getSnapshot: 'config:get-settings-snapshot',
  savePatch: 'config:save-settings-patch',
  openFile: 'config:open-settings-file',
  openWindow: 'config:open-settings-window',
} as const;

const configSettingsAPI: ConfigSettingsAPI = {
  getSnapshot: (): Promise<ConfigSettingsSnapshot> =>
    ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.getSnapshot),
  savePatch: (patch: ConfigSettingsPatch): Promise<ConfigSettingsSaveResult> =>
    ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.savePatch, patch),
  openSettingsFile: (): Promise<boolean> => ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.openFile),
  openSettingsWindow: (): Promise<boolean> => ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.openWindow),
};

contextBridge.exposeInMainWorld('configSettingsAPI', configSettingsAPI);
