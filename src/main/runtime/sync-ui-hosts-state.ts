import {
  readSyncHostsState,
  writeSyncHostsState,
  type SyncHostsState,
} from '../../shared/sync/sync-hosts-store';

interface SyncUiHostsStateDeps {
  hostsFilePath: string;
  broadcastStateChanged: () => void;
}

export function createSyncUiHostsState(deps: SyncUiHostsStateDeps) {
  function readState(): SyncHostsState {
    return readSyncHostsState(deps.hostsFilePath);
  }

  function writeState(state: SyncHostsState): void {
    writeSyncHostsState(deps.hostsFilePath, state);
    deps.broadcastStateChanged();
  }

  return { readState, writeState };
}
