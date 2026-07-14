import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import {
  isValidSyncHost,
  upsertSyncHost,
  type SyncDirection,
} from '../../shared/sync/sync-hosts-store';
import type { SyncUiRunRequest, SyncUiSnapshot } from '../../types/sync-ui';
import { createSyncUiHostsState } from './sync-ui-hosts-state';
import { registerSyncUiIpcHandlers } from './sync-ui-ipc-handlers';
import { createSyncUiRunLifecycle } from './sync-ui-run-lifecycle';
import { runSyncLauncher } from './sync-launcher-client';
import { createSyncUiSnapshots } from './sync-ui-snapshots';

interface SyncUiWindowLike {
  isDestroyed(): boolean;
  webContents: { send(channel: string, payload?: unknown): void };
}

export interface SyncUiRuntimeDeps {
  ipcMain: {
    handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): unknown;
  };
  hostsFilePath: string;
  snapshotsDir: string;
  getDbPath: () => string;
  resolveLauncherCommand: () => string[];
  runLauncher: typeof runSyncLauncher;
  getWindow: () => SyncUiWindowLike | null;
  pickSnapshotFile: () => Promise<string | null>;
  revealPath: (targetPath: string) => void;
  nowMs: () => number;
  log?: (message: string) => void;
  notify?: (payload: { title: string; body: string; variant: 'success' | 'error' }) => void;
}

export function createSyncUiRuntime(deps: SyncUiRuntimeDeps) {
  function sendToWindow(channel: string, payload?: unknown): void {
    const window = deps.getWindow();
    if (!window || window.isDestroyed()) return;
    window.webContents.send(channel, payload);
  }

  function broadcastStateChanged(): void {
    sendToWindow(IPC_CHANNELS.event.syncUiStateChanged);
  }

  const hostsState = createSyncUiHostsState({
    hostsFilePath: deps.hostsFilePath,
    broadcastStateChanged,
  });
  const runLifecycle = createSyncUiRunLifecycle({
    runLauncher: deps.runLauncher,
    resolveLauncherCommand: deps.resolveLauncherCommand,
    log: deps.log,
    notify: deps.notify,
    sendToWindow,
    broadcastStateChanged,
  });
  const snapshots = createSyncUiSnapshots({
    snapshotsDir: deps.snapshotsDir,
    nowMs: deps.nowMs,
    startRun: runLifecycle.startRun,
    broadcastStateChanged,
  });

  function getSnapshot(): SyncUiSnapshot {
    return {
      dbPath: deps.getDbPath(),
      hosts: hostsState.readState(),
      snapshotsDir: deps.snapshotsDir,
      snapshots: snapshots.listSnapshots(),
      run: runLifecycle.runState(),
    };
  }

  function runHostSync(request: SyncUiRunRequest, options: { notify?: boolean } = {}) {
    const host = request.host.trim();
    if (!isValidSyncHost(host)) {
      return { started: false, runId: null, reason: `Invalid sync host: ${request.host}` };
    }
    const direction: SyncDirection = request.direction ?? 'both';
    const args = ['sync', host];
    if (direction === 'push') args.push('--push');
    if (direction === 'pull') args.push('--pull');
    // The sync window runs inside the resident app, so waiting for all app
    // writers to disappear would make manual and automatic sync impossible.
    // WAL keeps snapshots consistent, transactional merges serialize with live
    // writes, and unfinished sessions wait for a later sync after finalization.
    args.push('--force', '--json');
    const result = runLifecycle.startRun('host-sync', host, args, options);
    if (result.started) {
      // Persist before launcher bookkeeping lands so the host appears immediately.
      hostsState.writeState(
        upsertSyncHost(
          hostsState.readState(),
          { host, direction: request.direction },
          deps.nowMs(),
        ),
      );
    }
    return result;
  }

  function registerHandlers(): void {
    registerSyncUiIpcHandlers({
      ipcMain: deps.ipcMain,
      nowMs: deps.nowMs,
      revealPath: deps.revealPath,
      pickSnapshotFile: deps.pickSnapshotFile,
      getSnapshot,
      readState: hostsState.readState,
      writeState: hostsState.writeState,
      runHostSync,
      cancelRun: runLifecycle.cancelRun,
      checkHost: runLifecycle.checkHost,
      createSnapshot: snapshots.createSnapshot,
      mergeSnapshotFile: snapshots.mergeSnapshotFile,
      deleteSnapshot: snapshots.deleteSnapshot,
    });
  }

  return {
    registerHandlers,
    getSnapshot,
    readState: hostsState.readState,
    runHostSync,
    checkHost: runLifecycle.checkHost,
    cancelRun: runLifecycle.cancelRun,
    shutdown: runLifecycle.shutdown,
    isRunning: runLifecycle.isRunning,
  };
}
