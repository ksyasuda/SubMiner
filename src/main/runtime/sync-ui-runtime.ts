import fs from 'node:fs';
import path from 'node:path';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import {
  isValidSyncHost,
  readSyncHostsState,
  removeSyncHost,
  upsertSyncHost,
  writeSyncHostsState,
  type SyncDirection,
  type SyncHostsState,
} from '../../shared/sync/sync-hosts-store';
import type { SyncProgressEvent } from '../../shared/sync/sync-events';
import type {
  SyncUiCheckResult,
  SyncUiHostUpdateRequest,
  SyncUiRunKind,
  SyncUiRunRequest,
  SyncUiRunState,
  SyncUiSnapshot,
  SyncUiSnapshotFile,
  SyncUiStartResult,
} from '../../types/sync-ui';
import type {
  SyncLauncherResolution,
  SyncLauncherRunHandle,
  SyncLauncherRunResult,
} from './sync-launcher-client';
import { runSyncLauncher } from './sync-launcher-client';

export interface SyncUiWindowLike {
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
  resolveLauncherCommand: () => SyncLauncherResolution;
  runLauncher: typeof runSyncLauncher;
  getWindow: () => SyncUiWindowLike | null;
  pickSnapshotFile: () => Promise<string | null>;
  revealPath: (targetPath: string) => void;
  nowMs: () => number;
  log?: (message: string) => void;
  notify?: (payload: { title: string; body: string; variant: 'success' | 'error' }) => void;
}

interface ActiveRun {
  id: number;
  kind: SyncUiRunKind;
  host: string | null;
  handle: SyncLauncherRunHandle;
  resultSeen: boolean;
}

function formatSnapshotName(nowMs: number): string {
  const iso = new Date(nowMs).toISOString();
  const stamp = iso.slice(0, 19).replace(/[-:]/g, '').replace('T', '-');
  return `immersion-${stamp}.sqlite`;
}

export function createSyncUiRuntime(deps: SyncUiRuntimeDeps) {
  let runCounter = 0;
  let currentRun: ActiveRun | null = null;

  function sendToWindow(channel: string, payload?: unknown): void {
    const window = deps.getWindow();
    if (!window || window.isDestroyed()) return;
    window.webContents.send(channel, payload);
  }

  function broadcastStateChanged(): void {
    sendToWindow(IPC_CHANNELS.event.syncUiStateChanged);
  }

  function readState(): SyncHostsState {
    return readSyncHostsState(deps.hostsFilePath);
  }

  function writeState(state: SyncHostsState): SyncHostsState {
    writeSyncHostsState(deps.hostsFilePath, state);
    broadcastStateChanged();
    return state;
  }

  function listSnapshots(): SyncUiSnapshotFile[] {
    let names: string[];
    try {
      names = fs.readdirSync(deps.snapshotsDir);
    } catch {
      return [];
    }
    const files: SyncUiSnapshotFile[] = [];
    for (const name of names) {
      if (!name.endsWith('.sqlite')) continue;
      const filePath = path.join(deps.snapshotsDir, name);
      try {
        const stat = fs.statSync(filePath);
        if (!stat.isFile()) continue;
        files.push({
          path: filePath,
          name,
          sizeBytes: stat.size,
          modifiedAtMs: stat.mtimeMs,
        });
      } catch {
        // file disappeared between readdir and stat
      }
    }
    files.sort((a, b) => b.modifiedAtMs - a.modifiedAtMs);
    return files;
  }

  function runState(): SyncUiRunState {
    return currentRun
      ? { running: true, runId: currentRun.id, kind: currentRun.kind, host: currentRun.host }
      : { running: false, runId: null, kind: null, host: null };
  }

  function getSnapshot(): SyncUiSnapshot {
    const resolution = deps.resolveLauncherCommand();
    return {
      dbPath: deps.getDbPath(),
      hosts: readState(),
      snapshotsDir: deps.snapshotsDir,
      snapshots: listSnapshots(),
      run: runState(),
      launcherPath: resolution.command ? resolution.command.join(' ') : null,
    };
  }

  function startRun(
    kind: SyncUiRunKind,
    host: string | null,
    args: string[],
    options: { notify?: boolean } = {},
  ): SyncUiStartResult {
    if (currentRun) {
      return {
        started: false,
        runId: null,
        reason: 'A sync operation is already running.',
      };
    }
    const resolution = deps.resolveLauncherCommand();
    if (!resolution.command) {
      return { started: false, runId: null, reason: resolution.error };
    }
    runCounter += 1;
    const runId = runCounter;
    const run: ActiveRun = {
      id: runId,
      kind,
      host,
      resultSeen: false,
      handle: deps.runLauncher({
        command: resolution.command,
        args,
        onEvent: (event: SyncProgressEvent) => {
          if (event.type === 'result') run.resultSeen = true;
          sendToWindow(IPC_CHANNELS.event.syncUiProgress, { runId, kind, host, event });
        },
        onStderr: (text) => deps.log?.(`[sync-ui] ${text.trimEnd()}`),
      }),
    };
    currentRun = run;
    void run.handle.done.then((result: SyncLauncherRunResult) => {
      if (currentRun?.id === runId) currentRun = null;
      // If the launcher died without emitting a result event (spawn failure,
      // kill), synthesize one so the renderer can settle its progress view.
      if (!run.resultSeen) {
        sendToWindow(IPC_CHANNELS.event.syncUiProgress, {
          runId,
          kind,
          host,
          event: { type: 'result', ok: result.ok, error: result.error },
        });
      }
      broadcastStateChanged();
      if (options.notify && deps.notify) {
        const target = host ?? 'local database';
        deps.notify(
          result.ok
            ? { title: 'Sync complete', body: `Synced with ${target}`, variant: 'success' }
            : {
                title: 'Sync failed',
                body: `${target}: ${result.error ?? 'unknown error'}`,
                variant: 'error',
              },
        );
      }
    });
    return { started: true, runId, reason: null };
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
    if (request.force) args.push('--force');
    args.push('--json');
    const result = startRun('host-sync', host, args, options);
    if (result.started) {
      // Remember the host (and the direction used) before the launcher's own
      // bookkeeping lands, so it shows up in the UI immediately.
      writeState(upsertSyncHost(readState(), { host, direction: request.direction }, deps.nowMs()));
    }
    return result;
  }

  function checkHost(host: string): Promise<SyncUiCheckResult> {
    const trimmed = host.trim();
    const failed = (error: string): SyncUiCheckResult => ({
      host: trimmed,
      sshOk: false,
      remoteCommand: null,
      remoteVersion: null,
      ok: false,
      error,
    });
    if (!isValidSyncHost(trimmed)) {
      return Promise.resolve(failed(`Invalid sync host: ${host}`));
    }
    const resolution = deps.resolveLauncherCommand();
    if (!resolution.command) {
      return Promise.resolve(failed(resolution.error ?? 'Launcher unavailable.'));
    }
    return new Promise((resolve) => {
      let checkResult: SyncUiCheckResult | null = null;
      const handle = deps.runLauncher({
        command: resolution.command!,
        args: ['sync', trimmed, '--check', '--json'],
        onEvent: (event) => {
          if (event.type === 'check-result') {
            checkResult = {
              host: event.host,
              sshOk: event.sshOk,
              remoteCommand: event.remoteCommand,
              remoteVersion: event.remoteVersion,
              ok: event.ok,
              error: event.error,
            };
          }
        },
        onStderr: (text) => deps.log?.(`[sync-ui check] ${text.trimEnd()}`),
      });
      void handle.done.then((result) => {
        resolve(checkResult ?? failed(result.error ?? 'Connection check failed.'));
      });
    });
  }

  function createSnapshot(): SyncUiStartResult {
    const outPath = path.join(deps.snapshotsDir, formatSnapshotName(deps.nowMs()));
    return startRun('snapshot', null, ['sync', '--snapshot', outPath, '--json']);
  }

  function mergeSnapshotFile(filePath: string, force = false): SyncUiStartResult {
    if (!fs.existsSync(filePath)) {
      return { started: false, runId: null, reason: `Snapshot file not found: ${filePath}` };
    }
    const args = ['sync', '--merge', filePath];
    if (force) args.push('--force');
    args.push('--json');
    return startRun('merge', null, args);
  }

  function deleteSnapshot(filePath: string): SyncUiSnapshotFile[] {
    const resolved = path.resolve(filePath);
    const dir = path.resolve(deps.snapshotsDir);
    if (resolved !== path.join(dir, path.basename(resolved)) || !resolved.endsWith('.sqlite')) {
      throw new Error('Refusing to delete a file outside the snapshots directory.');
    }
    fs.rmSync(resolved, { force: true });
    broadcastStateChanged();
    return listSnapshots();
  }

  function cancelRun(): boolean {
    if (!currentRun) return false;
    currentRun.handle.cancel();
    return true;
  }

  function registerHandlers(): void {
    const channels = IPC_CHANNELS.request;
    deps.ipcMain.handle(channels.syncUiGetSnapshot, () => getSnapshot());
    deps.ipcMain.handle(channels.syncUiSaveHost, (_event, update) =>
      writeState(upsertSyncHost(readState(), update as SyncUiHostUpdateRequest, deps.nowMs())),
    );
    deps.ipcMain.handle(channels.syncUiRemoveHost, (_event, host) =>
      writeState(removeSyncHost(readState(), String(host))),
    );
    deps.ipcMain.handle(channels.syncUiSetAutoSyncInterval, (_event, minutes) => {
      const value = Number(minutes);
      if (!Number.isFinite(value) || value < 1 || value > 24 * 60) {
        throw new Error('Auto-sync interval must be between 1 and 1440 minutes.');
      }
      return writeState({ ...readState(), autoSyncIntervalMinutes: Math.floor(value) });
    });
    deps.ipcMain.handle(channels.syncUiRunSync, (_event, request) =>
      runHostSync(request as SyncUiRunRequest),
    );
    deps.ipcMain.handle(channels.syncUiCancelRun, () => cancelRun());
    deps.ipcMain.handle(channels.syncUiCheckHost, (_event, host) => checkHost(String(host)));
    deps.ipcMain.handle(channels.syncUiCreateSnapshot, () => createSnapshot());
    deps.ipcMain.handle(channels.syncUiMergeSnapshotFile, (_event, filePath, force) =>
      mergeSnapshotFile(String(filePath), force === true),
    );
    deps.ipcMain.handle(channels.syncUiDeleteSnapshot, (_event, filePath) =>
      deleteSnapshot(String(filePath)),
    );
    deps.ipcMain.handle(channels.syncUiRevealSnapshot, (_event, filePath) => {
      deps.revealPath(String(filePath));
      return true;
    });
    deps.ipcMain.handle(channels.syncUiPickSnapshotFile, () => deps.pickSnapshotFile());
  }

  return {
    registerHandlers,
    getSnapshot,
    readState,
    runHostSync,
    checkHost,
    cancelRun,
    isRunning: () => currentRun !== null,
  };
}
