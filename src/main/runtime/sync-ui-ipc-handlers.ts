import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { removeSyncHost, upsertSyncHost } from '../../shared/sync/sync-hosts-store';
import type {
  SyncUiHostUpdateRequest,
  SyncUiRunRequest,
  SyncUiSnapshot,
  SyncUiStartResult,
} from '../../types/sync-ui';

interface SyncUiIpcHandlersDeps {
  ipcMain: {
    handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): unknown;
  };
  nowMs: () => number;
  revealPath: (targetPath: string) => void;
  pickSnapshotFile: () => Promise<string | null>;
  getSnapshot: () => SyncUiSnapshot;
  readState: () => SyncUiSnapshot['hosts'];
  writeState: (state: SyncUiSnapshot['hosts']) => void;
  runHostSync: (request: SyncUiRunRequest) => SyncUiStartResult;
  cancelRun: () => boolean;
  checkHost: (host: string) => unknown;
  createSnapshot: () => SyncUiStartResult;
  mergeSnapshotFile: (filePath: string, force?: boolean) => SyncUiStartResult;
  deleteSnapshot: (filePath: string) => void;
}

export function registerSyncUiIpcHandlers(deps: SyncUiIpcHandlersDeps): void {
  const channels = IPC_CHANNELS.request;
  deps.ipcMain.handle(channels.syncUiGetSnapshot, () => deps.getSnapshot());
  deps.ipcMain.handle(channels.syncUiSaveHost, (_event, update) => {
    deps.writeState(
      upsertSyncHost(deps.readState(), update as SyncUiHostUpdateRequest, deps.nowMs()),
    );
  });
  deps.ipcMain.handle(channels.syncUiRemoveHost, (_event, host) => {
    deps.writeState(removeSyncHost(deps.readState(), String(host)));
  });
  deps.ipcMain.handle(channels.syncUiSetAutoSyncInterval, (_event, minutes) => {
    const value = Number(minutes);
    if (!Number.isFinite(value) || value < 1 || value > 24 * 60) {
      throw new Error('Auto-sync interval must be between 1 and 1440 minutes.');
    }
    deps.writeState({ ...deps.readState(), autoSyncIntervalMinutes: Math.floor(value) });
  });
  deps.ipcMain.handle(channels.syncUiRunSync, (_event, request) =>
    deps.runHostSync(request as SyncUiRunRequest),
  );
  deps.ipcMain.handle(channels.syncUiCancelRun, () => deps.cancelRun());
  deps.ipcMain.handle(channels.syncUiCheckHost, (_event, host) => deps.checkHost(String(host)));
  deps.ipcMain.handle(channels.syncUiCreateSnapshot, () => deps.createSnapshot());
  deps.ipcMain.handle(channels.syncUiMergeSnapshotFile, (_event, filePath, force) =>
    deps.mergeSnapshotFile(String(filePath), force === true),
  );
  deps.ipcMain.handle(channels.syncUiDeleteSnapshot, (_event, filePath) => {
    deps.deleteSnapshot(String(filePath));
  });
  deps.ipcMain.handle(channels.syncUiRevealSnapshot, (_event, filePath) => {
    deps.revealPath(String(filePath));
    return true;
  });
  deps.ipcMain.handle(channels.syncUiPickSnapshotFile, () => deps.pickSnapshotFile());
}
