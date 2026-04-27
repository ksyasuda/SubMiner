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

export function createEnsureStatsServerUrlHandler(deps: EnsureStatsServerUrlDeps): () => string {
  return () => {
    const state = deps.readBackgroundState();
    if (!state) {
      deps.removeBackgroundState();
    } else if (state.pid === deps.currentPid && !deps.hasLocalStatsServer()) {
      deps.removeBackgroundState();
    } else if (!deps.isProcessAlive(state.pid)) {
      deps.removeBackgroundState();
    } else if (state.pid !== deps.currentPid) {
      return formatStatsServerUrl(state.port);
    }

    if (!deps.hasLocalStatsServer()) {
      deps.startLocalStatsServer();
    }
    return formatStatsServerUrl(deps.getConfiguredPort());
  };
}
