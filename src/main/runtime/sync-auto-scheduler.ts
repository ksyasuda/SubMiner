import type { SyncDirection, SyncHostsState } from '../../shared/sync/sync-hosts-store';

export interface SyncAutoSchedulerDeps {
  readState: () => SyncHostsState;
  isRunning: () => boolean;
  triggerHostSync: (host: string, direction: SyncDirection) => void;
  nowMs: () => number;
  log?: (message: string) => void;
  tickIntervalMs?: number;
  setIntervalFn?: typeof setInterval;
  clearIntervalFn?: typeof clearInterval;
}

const DEFAULT_TICK_INTERVAL_MS = 60_000;

export function createSyncAutoScheduler(deps: SyncAutoSchedulerDeps) {
  let timer: NodeJS.Timeout | null = null;

  function tick(): void {
    if (deps.isRunning()) return;
    const state = deps.readState();
    const intervalMs = state.autoSyncIntervalMinutes * 60_000;
    const now = deps.nowMs();
    // A failed sync still updates lastSyncAtMs, so errors retry on the next
    // interval instead of every tick.
    const due = state.hosts.find(
      (entry) =>
        entry.autoSync && (entry.lastSyncAtMs === null || now - entry.lastSyncAtMs >= intervalMs),
    );
    if (!due) return;
    deps.log?.(`Auto-sync starting for ${due.host} (${due.direction})`);
    try {
      deps.triggerHostSync(due.host, due.direction);
    } catch (error) {
      deps.log?.(
        `Auto-sync failed to start: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }

  function start(): void {
    if (timer !== null) return;
    const setIntervalFn = deps.setIntervalFn ?? setInterval;
    timer = setIntervalFn(tick, deps.tickIntervalMs ?? DEFAULT_TICK_INTERVAL_MS);
  }

  function stop(): void {
    if (timer === null) return;
    const clearIntervalFn = deps.clearIntervalFn ?? clearInterval;
    clearIntervalFn(timer);
    timer = null;
  }

  return { tick, start, stop };
}
