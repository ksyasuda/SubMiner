import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { shell } from 'electron';
import { sanitizeStartupEnv } from './main-entry-runtime';
import {
  isBackgroundStatsServerProcessAlive,
  readBackgroundStatsServerState,
  removeBackgroundStatsServerState,
  resolveBackgroundStatsServerUrl,
} from './main/runtime/stats-daemon';
import {
  createRunStatsDaemonControlHandler,
  type StatsDaemonControlArgs,
} from './stats-daemon-control';
import {
  type StatsCliCommandResponse,
  writeStatsCliCommandResponse,
} from './main/runtime/stats-cli-command';

const STATS_DAEMON_RESPONSE_TIMEOUT_MS = 12_000;

function readFlagValue(argv: string[], flag: string): string | undefined {
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg) continue;
    if (arg === flag) {
      const value = argv[i + 1];
      if (value && !value.startsWith('--')) {
        return value;
      }
      return undefined;
    }
    if (arg.startsWith(`${flag}=`)) {
      return arg.split('=', 2)[1];
    }
  }
  return undefined;
}

function hasFlag(argv: string[], flag: string): boolean {
  return argv.includes(flag);
}

function parseControlArgs(argv: string[], userDataPath: string): StatsDaemonControlArgs {
  return {
    action: hasFlag(argv, '--stats-daemon-stop') ? 'stop' : 'start',
    responsePath: readFlagValue(argv, '--stats-response-path'),
    openBrowser: hasFlag(argv, '--stats-daemon-open-browser'),
    daemonScriptPath: path.join(__dirname, 'stats-daemon-runner.js'),
    userDataPath,
  };
}

async function waitForDaemonResponse(responsePath: string): Promise<StatsCliCommandResponse> {
  const deadline = Date.now() + STATS_DAEMON_RESPONSE_TIMEOUT_MS;
  while (Date.now() < deadline) {
    try {
      if (fs.existsSync(responsePath)) {
        return JSON.parse(fs.readFileSync(responsePath, 'utf8')) as StatsCliCommandResponse;
      }
    } catch {
      // retry until timeout
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }

  return {
    ok: false,
    error: 'Timed out waiting for stats daemon startup response.',
  };
}

export async function runStatsDaemonControlFromProcess(userDataPath: string): Promise<number> {
  const args = parseControlArgs(process.argv, userDataPath);
  const statePath = path.join(userDataPath, 'stats-daemon.json');

  const writeFailureResponse = (message: string): void => {
    if (args.responsePath) {
      try {
        writeStatsCliCommandResponse(args.responsePath, {
          ok: false,
          error: message,
        });
      } catch {
        // ignore secondary response-write failures
      }
    }
  };

  const handler = createRunStatsDaemonControlHandler({
    statePath,
    readState: () => readBackgroundStatsServerState(statePath),
    removeState: () => {
      removeBackgroundStatsServerState(statePath);
    },
    isProcessAlive: (pid) => isBackgroundStatsServerProcessAlive(pid),
    resolveUrl: (state) => resolveBackgroundStatsServerUrl(state),
    spawnDaemon: async (options) => {
      const childArgs = [options.scriptPath, '--stats-user-data-path', options.userDataPath];
      if (options.responsePath) {
        childArgs.push('--stats-response-path', options.responsePath);
      }
      const logLevel = readFlagValue(process.argv, '--log-level');
      if (logLevel) {
        childArgs.push('--log-level', logLevel);
      }
      const child = spawn(process.execPath, childArgs, {
        detached: true,
        stdio: 'ignore',
        env: {
          ...sanitizeStartupEnv(process.env),
          ELECTRON_RUN_AS_NODE: '1',
        },
      });
      child.unref();
      return child.pid ?? 0;
    },
    waitForDaemonResponse,
    openExternal: async (url) => shell.openExternal(url),
    writeResponse: writeStatsCliCommandResponse,
    killProcess: (pid, signal) => {
      process.kill(pid, signal);
    },
    sleep: async (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
  });

  try {
    return await handler(args);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    writeFailureResponse(message);
    return 1;
  }
}
