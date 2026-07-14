import { contextBridge, ipcRenderer } from 'electron';
import { IPC_CHANNELS } from './shared/ipc/contracts';
import type {
  SyncUiAPI,
  SyncUiCheckResult,
  SyncUiHostUpdateRequest,
  SyncUiProgressPayload,
  SyncUiRunRequest,
  SyncUiSnapshot,
  SyncUiStartResult,
} from './types/sync-ui';

const syncUiAPI: SyncUiAPI = {
  getSnapshot: (): Promise<SyncUiSnapshot> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiGetSnapshot),
  saveHost: (update: SyncUiHostUpdateRequest): Promise<void> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiSaveHost, update),
  removeHost: (host: string): Promise<void> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiRemoveHost, host),
  setAutoSyncInterval: (minutes: number): Promise<void> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiSetAutoSyncInterval, minutes),
  runSync: (request: SyncUiRunRequest): Promise<SyncUiStartResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiRunSync, request),
  cancelRun: (): Promise<boolean> => ipcRenderer.invoke(IPC_CHANNELS.request.syncUiCancelRun),
  checkHost: (host: string): Promise<SyncUiCheckResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiCheckHost, host),
  createSnapshot: (): Promise<SyncUiStartResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiCreateSnapshot),
  mergeSnapshotFile: (path: string, force?: boolean): Promise<SyncUiStartResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiMergeSnapshotFile, path, force === true),
  deleteSnapshot: (path: string): Promise<void> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiDeleteSnapshot, path),
  revealSnapshot: (path: string): Promise<boolean> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiRevealSnapshot, path),
  pickSnapshotFile: (): Promise<string | null> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.syncUiPickSnapshotFile),
  onProgress: (listener: (payload: SyncUiProgressPayload) => void): (() => void) => {
    const handler = (_event: unknown, payload: SyncUiProgressPayload): void => listener(payload);
    ipcRenderer.on(IPC_CHANNELS.event.syncUiProgress, handler);
    return () => ipcRenderer.removeListener(IPC_CHANNELS.event.syncUiProgress, handler);
  },
  onStateChanged: (listener: () => void): (() => void) => {
    const handler = (): void => listener();
    ipcRenderer.on(IPC_CHANNELS.event.syncUiStateChanged, handler);
    return () => ipcRenderer.removeListener(IPC_CHANNELS.event.syncUiStateChanged, handler);
  },
};

contextBridge.exposeInMainWorld('syncUiAPI', syncUiAPI);
