import type { BackgroundStatsServerState } from './stats-daemon';

type EnsureStatsServerUrlDeps = {
  currentPid: number;
  readBackgroundState: () => BackgroundStatsServerState | null;
  removeBackgroundState: () => void;
  isProcessAlive: (pid: number) => boolean;
  hasLocalStatsServer: () => boolean;
  startLocalStatsServer: () => void;
  getConfiguredPort: () => number;
};

function formatStatsServerUrl(port: number): string {
  return `http://127.0.0.1:${port}`;
}

export type EnsureStatsServerUrlResult = { url: string; source: 'background' | 'local' };

export function createEnsureStatsServerUrlHandler(
  deps: EnsureStatsServerUrlDeps,
): () => EnsureStatsServerUrlResult {
  return () => {
    const state = deps.readBackgroundState();
    if (!state) {
      deps.removeBackgroundState();
    } else if (state.pid === deps.currentPid && !deps.hasLocalStatsServer()) {
      deps.removeBackgroundState();
    } else if (!deps.isProcessAlive(state.pid)) {
      deps.removeBackgroundState();
    } else if (state.pid !== deps.currentPid) {
      return { url: formatStatsServerUrl(state.port), source: 'background' };
    }

    if (!deps.hasLocalStatsServer()) {
      deps.startLocalStatsServer();
    }
    return { url: formatStatsServerUrl(deps.getConfiguredPort()), source: 'local' };
  };
}
