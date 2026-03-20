import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { runAppCommandAttached } from '../mpv.js';
import { sleep } from '../util.js';
import type { LauncherCommandContext } from './context.js';

type StatsCommandResponse = {
  ok: boolean;
  url?: string;
  error?: string;
};

type StatsCommandDeps = {
  createTempDir: (prefix: string) => string;
  joinPath: (...parts: string[]) => string;
  runAppCommandAttached: (
    appPath: string,
    appArgs: string[],
    logLevel: LauncherCommandContext['args']['logLevel'],
    label: string,
  ) => Promise<number>;
  waitForStatsResponse: (
    responsePath: string,
    signal?: AbortSignal,
  ) => Promise<StatsCommandResponse>;
  removeDir: (targetPath: string) => void;
};

const STATS_STARTUP_RESPONSE_TIMEOUT_MS = 12_000;

type StatsResponseWait = {
  controller: AbortController;
  promise: Promise<{ kind: 'response'; response: StatsCommandResponse }>;
};

const defaultDeps: StatsCommandDeps = {
  createTempDir: (prefix) => fs.mkdtempSync(path.join(os.tmpdir(), prefix)),
  joinPath: (...parts) => path.join(...parts),
  runAppCommandAttached: (appPath, appArgs, logLevel, label) =>
    runAppCommandAttached(appPath, appArgs, logLevel, label),
  waitForStatsResponse: async (responsePath, signal) => {
    const deadline = Date.now() + STATS_STARTUP_RESPONSE_TIMEOUT_MS;
    while (Date.now() < deadline) {
      if (signal?.aborted) {
        return {
          ok: false,
          error: 'Cancelled waiting for stats dashboard startup response.',
        };
      }
      try {
        if (fs.existsSync(responsePath)) {
          return JSON.parse(fs.readFileSync(responsePath, 'utf8')) as StatsCommandResponse;
        }
      } catch {
        // retry until timeout
      }
      await sleep(100);
    }
    return {
      ok: false,
      error: 'Timed out waiting for stats dashboard startup response.',
    };
  },
  removeDir: (targetPath) => {
    fs.rmSync(targetPath, { recursive: true, force: true });
  },
};

async function performStartupHandshake(
  createResponseWait: () => StatsResponseWait,
  attachedExitPromise: Promise<number>,
): Promise<boolean> {
  const responseWait = createResponseWait();
  const startupResult = await Promise.race([
    responseWait.promise,
    attachedExitPromise.then((status) => ({ kind: 'exit' as const, status })),
  ]);

  if (startupResult.kind === 'exit') {
    if (startupResult.status !== 0) {
      responseWait.controller.abort();
      throw new Error(`Stats app exited before startup response (status ${startupResult.status}).`);
    }

    const response = await responseWait.promise.then((result) => result.response);
    if (!response.ok) {
      throw new Error(response.error || 'Stats dashboard failed to start.');
    }
    return true;
  }

  if (!startupResult.response.ok) {
    throw new Error(startupResult.response.error || 'Stats dashboard failed to start.');
  }

  const exitStatus = await attachedExitPromise;
  if (exitStatus !== 0) {
    throw new Error(`Stats app exited with status ${exitStatus}.`);
  }

  return true;
}

export async function runStatsCommand(
  context: LauncherCommandContext,
  deps: Partial<StatsCommandDeps> = {},
): Promise<boolean> {
  const resolvedDeps: StatsCommandDeps = { ...defaultDeps, ...deps };
  const { args, appPath } = context;
  if (!args.stats || !appPath) {
    return false;
  }

  const tempDir = resolvedDeps.createTempDir('subminer-stats-');
  const responsePath = resolvedDeps.joinPath(tempDir, 'response.json');

  const createResponseWait = () => {
    const controller = new AbortController();
    return {
      controller,
      promise: resolvedDeps
        .waitForStatsResponse(responsePath, controller.signal)
        .then((response) => ({ kind: 'response' as const, response })),
    };
  };

  try {
    const forwarded = args.statsCleanup
      ? ['--stats', '--stats-response-path', responsePath]
      : args.statsStop
        ? ['--stats-daemon-stop', '--stats-response-path', responsePath]
        : args.statsBackground
          ? ['--stats-daemon-start', '--stats-response-path', responsePath]
          : ['--stats', '--stats-response-path', responsePath];
    if (args.statsCleanup) {
      forwarded.push('--stats-cleanup');
    }
    if (args.statsCleanupVocab) {
      forwarded.push('--stats-cleanup-vocab');
    }
    if (args.statsCleanupLifetime) {
      forwarded.push('--stats-cleanup-lifetime');
    }
    if (args.logLevel !== 'info') {
      forwarded.push('--log-level', args.logLevel);
    }
    const attachedExitPromise = resolvedDeps.runAppCommandAttached(
      appPath,
      forwarded,
      args.logLevel,
      'stats',
    );

    if (args.statsStop) {
      const status = await attachedExitPromise;
      if (status !== 0) {
        throw new Error(`Stats app exited with status ${status}.`);
      }
      return true;
    }

    return await performStartupHandshake(createResponseWait, attachedExitPromise);
  } finally {
    resolvedDeps.removeDir(tempDir);
  }
}
