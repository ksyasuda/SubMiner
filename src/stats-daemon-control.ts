import type { BackgroundStatsServerState } from './main/runtime/stats-daemon';
import type { StatsCliCommandResponse } from './main/runtime/stats-cli-command';

export type StatsDaemonControlAction = 'start' | 'stop';

export type StatsDaemonControlArgs = {
  action: StatsDaemonControlAction;
  responsePath?: string;
  openBrowser: boolean;
  daemonScriptPath: string;
  userDataPath: string;
};

type SpawnStatsDaemonOptions = {
  scriptPath: string;
  responsePath?: string;
  userDataPath: string;
};

export function createRunStatsDaemonControlHandler(deps: {
  statePath: string;
  readState: () => BackgroundStatsServerState | null;
  removeState: () => void;
  isProcessAlive: (pid: number) => boolean;
  resolveUrl: (state: Pick<BackgroundStatsServerState, 'port'>) => string;
  spawnDaemon: (options: SpawnStatsDaemonOptions) => Promise<number> | number;
  waitForDaemonResponse: (responsePath: string) => Promise<StatsCliCommandResponse>;
  openExternal: (url: string) => Promise<unknown>;
  writeResponse: (responsePath: string, payload: StatsCliCommandResponse) => void;
  killProcess: (pid: number, signal: NodeJS.Signals) => void;
  sleep: (ms: number) => Promise<void>;
}) {
  const writeResponseSafe = (
    responsePath: string | undefined,
    payload: StatsCliCommandResponse,
  ): void => {
    if (!responsePath) return;
    deps.writeResponse(responsePath, payload);
  };

  return async (args: StatsDaemonControlArgs): Promise<number> => {
    if (args.action === 'start') {
      const state = deps.readState();
      if (state) {
        if (deps.isProcessAlive(state.pid)) {
          const url = deps.resolveUrl(state);
          writeResponseSafe(args.responsePath, { ok: true, url });
          if (args.openBrowser) {
            await deps.openExternal(url);
          }
          return 0;
        }
        deps.removeState();
      }

      if (!args.responsePath) {
        throw new Error('Missing --stats-response-path for stats daemon start.');
      }

      await deps.spawnDaemon({
        scriptPath: args.daemonScriptPath,
        responsePath: args.responsePath,
        userDataPath: args.userDataPath,
      });
      const response = await deps.waitForDaemonResponse(args.responsePath);
      if (response.ok && args.openBrowser && response.url) {
        await deps.openExternal(response.url);
      }
      return response.ok ? 0 : 1;
    }

    const state = deps.readState();
    if (!state) {
      deps.removeState();
      writeResponseSafe(args.responsePath, { ok: true });
      return 0;
    }

    if (!deps.isProcessAlive(state.pid)) {
      deps.removeState();
      writeResponseSafe(args.responsePath, { ok: true });
      return 0;
    }

    deps.killProcess(state.pid, 'SIGTERM');
    const deadline = Date.now() + 2_000;
    while (Date.now() < deadline) {
      if (!deps.isProcessAlive(state.pid)) {
        deps.removeState();
        writeResponseSafe(args.responsePath, { ok: true });
        return 0;
      }
      await deps.sleep(50);
    }

    writeResponseSafe(args.responsePath, {
      ok: false,
      error: 'Timed out stopping background stats server.',
    });
    return 1;
  };
}
