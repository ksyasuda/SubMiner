import { contextBridge, ipcRenderer } from 'electron';
import type {
  SyncHostsState,
  SyncUiAPI,
  SyncUiCheckResult,
  SyncUiHostUpdateRequest,
  SyncUiProgressPayload,
  SyncUiRunRequest,
  SyncUiSnapshot,
  SyncUiSnapshotFile,
  SyncUiStartResult,
} from './types/sync-ui';

const SYNC_UI_IPC_CHANNELS = {
  getSnapshot: 'sync-ui:get-snapshot',
  saveHost: 'sync-ui:save-host',
  removeHost: 'sync-ui:remove-host',
  setAutoSyncInterval: 'sync-ui:set-auto-sync-interval',
  runSync: 'sync-ui:run-sync',
  cancelRun: 'sync-ui:cancel-run',
  checkHost: 'sync-ui:check-host',
  createSnapshot: 'sync-ui:create-snapshot',
  mergeSnapshotFile: 'sync-ui:merge-snapshot-file',
  deleteSnapshot: 'sync-ui:delete-snapshot',
  revealSnapshot: 'sync-ui:reveal-snapshot',
  pickSnapshotFile: 'sync-ui:pick-snapshot-file',
  progress: 'sync-ui:progress',
  stateChanged: 'sync-ui:state-changed',
} as const;

const syncUiAPI: SyncUiAPI = {
  getSnapshot: (): Promise<SyncUiSnapshot> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.getSnapshot),
  saveHost: (update: SyncUiHostUpdateRequest): Promise<SyncHostsState> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.saveHost, update),
  removeHost: (host: string): Promise<SyncHostsState> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.removeHost, host),
  setAutoSyncInterval: (minutes: number): Promise<SyncHostsState> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.setAutoSyncInterval, minutes),
  runSync: (request: SyncUiRunRequest): Promise<SyncUiStartResult> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.runSync, request),
  cancelRun: (): Promise<boolean> => ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.cancelRun),
  checkHost: (host: string): Promise<SyncUiCheckResult> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.checkHost, host),
  createSnapshot: (): Promise<SyncUiStartResult> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.createSnapshot),
  mergeSnapshotFile: (path: string): Promise<SyncUiStartResult> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.mergeSnapshotFile, path),
  deleteSnapshot: (path: string): Promise<SyncUiSnapshotFile[]> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.deleteSnapshot, path),
  revealSnapshot: (path: string): Promise<boolean> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.revealSnapshot, path),
  pickSnapshotFile: (): Promise<string | null> =>
    ipcRenderer.invoke(SYNC_UI_IPC_CHANNELS.pickSnapshotFile),
  onProgress: (listener: (payload: SyncUiProgressPayload) => void): (() => void) => {
    const handler = (_event: unknown, payload: SyncUiProgressPayload): void => listener(payload);
    ipcRenderer.on(SYNC_UI_IPC_CHANNELS.progress, handler);
    return () => ipcRenderer.removeListener(SYNC_UI_IPC_CHANNELS.progress, handler);
  },
  onStateChanged: (listener: () => void): (() => void) => {
    const handler = (): void => listener();
    ipcRenderer.on(SYNC_UI_IPC_CHANNELS.stateChanged, handler);
    return () => ipcRenderer.removeListener(SYNC_UI_IPC_CHANNELS.stateChanged, handler);
  },
};

contextBridge.exposeInMainWorld('syncUiAPI', syncUiAPI);
