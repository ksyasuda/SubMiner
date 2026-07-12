import type { SyncProgressEvent } from '../shared/sync/sync-events';
import type { SyncDirection, SyncHostsState } from '../shared/sync/sync-hosts-store';

export type { SyncProgressEvent } from '../shared/sync/sync-events';
export type { SyncDirection, SyncHostEntry, SyncHostsState } from '../shared/sync/sync-hosts-store';

export interface SyncUiSnapshotFile {
  path: string;
  name: string;
  sizeBytes: number;
  modifiedAtMs: number;
}

export type SyncUiRunKind = 'host-sync' | 'merge' | 'check' | 'snapshot';

export interface SyncUiRunState {
  running: boolean;
  runId: number | null;
  kind: SyncUiRunKind | null;
  host: string | null;
}

export interface SyncUiSnapshot {
  dbPath: string;
  hosts: SyncHostsState;
  snapshotsDir: string;
  snapshots: SyncUiSnapshotFile[];
  run: SyncUiRunState;
}

export interface SyncUiRunRequest {
  host: string;
  direction?: SyncDirection;
  force?: boolean;
}

export interface SyncUiStartResult {
  started: boolean;
  runId: number | null;
  reason: string | null;
}

export interface SyncUiCheckResult {
  host: string;
  sshOk: boolean;
  remoteCommand: string | null;
  remoteVersion: string | null;
  ok: boolean;
  error: string | null;
}

export interface SyncUiProgressPayload {
  runId: number;
  kind: SyncUiRunKind;
  host: string | null;
  event: SyncProgressEvent;
}

export interface SyncUiHostUpdateRequest {
  host: string;
  label?: string | null;
  direction?: SyncDirection;
  autoSync?: boolean;
}

export interface SyncUiAPI {
  getSnapshot: () => Promise<SyncUiSnapshot>;
  saveHost: (update: SyncUiHostUpdateRequest) => Promise<SyncHostsState>;
  removeHost: (host: string) => Promise<SyncHostsState>;
  setAutoSyncInterval: (minutes: number) => Promise<SyncHostsState>;
  runSync: (request: SyncUiRunRequest) => Promise<SyncUiStartResult>;
  cancelRun: () => Promise<boolean>;
  checkHost: (host: string) => Promise<SyncUiCheckResult>;
  createSnapshot: () => Promise<SyncUiStartResult>;
  mergeSnapshotFile: (path: string, force?: boolean) => Promise<SyncUiStartResult>;
  deleteSnapshot: (path: string) => Promise<SyncUiSnapshotFile[]>;
  revealSnapshot: (path: string) => Promise<boolean>;
  pickSnapshotFile: () => Promise<string | null>;
  onProgress: (listener: (payload: SyncUiProgressPayload) => void) => () => void;
  onStateChanged: (listener: () => void) => () => void;
}
