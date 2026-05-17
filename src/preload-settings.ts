import { contextBridge, ipcRenderer } from 'electron';
import type {
  ConfigSettingsAnkiListResult,
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
  getAnkiDeckNames: 'config-settings:anki-deck-names',
  getAnkiDeckFieldNames: 'config-settings:anki-deck-field-names',
  getAnkiModelNames: 'config-settings:anki-model-names',
  getAnkiModelFieldNames: 'config-settings:anki-model-field-names',
} as const;

const configSettingsAPI: ConfigSettingsAPI = {
  getSnapshot: (): Promise<ConfigSettingsSnapshot> =>
    ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.getSnapshot),
  savePatch: (patch: ConfigSettingsPatch): Promise<ConfigSettingsSaveResult> =>
    ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.savePatch, patch),
  openSettingsFile: (): Promise<boolean> => ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.openFile),
  openSettingsWindow: (): Promise<boolean> => ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.openWindow),
  getAnkiDeckNames: (draftUrl?: string): Promise<ConfigSettingsAnkiListResult> =>
    ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.getAnkiDeckNames, draftUrl),
  getAnkiDeckFieldNames: (
    deckName: string,
    draftUrl?: string,
  ): Promise<ConfigSettingsAnkiListResult> =>
    ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.getAnkiDeckFieldNames, deckName, draftUrl),
  getAnkiModelNames: (draftUrl?: string): Promise<ConfigSettingsAnkiListResult> =>
    ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.getAnkiModelNames, draftUrl),
  getAnkiModelFieldNames: (
    modelName: string,
    draftUrl?: string,
  ): Promise<ConfigSettingsAnkiListResult> =>
    ipcRenderer.invoke(SETTINGS_IPC_CHANNELS.getAnkiModelFieldNames, modelName, draftUrl),
};

contextBridge.exposeInMainWorld('configSettingsAPI', configSettingsAPI);
