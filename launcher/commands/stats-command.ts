import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { launchAppCommandDetached, runAppCommandAttached } from '../mpv.js';
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
  launchAppCommandDetached: (
    appPath: string,
    appArgs: string[],
    logLevel: LauncherCommandContext['args']['logLevel'],
    label: string,
  ) => void;
  waitForStatsResponse: (responsePath: string) => Promise<StatsCommandResponse>;
  removeDir: (targetPath: string) => void;
};

const STATS_STARTUP_RESPONSE_TIMEOUT_MS = 12_000;

const defaultDeps: StatsCommandDeps = {
  createTempDir: (prefix) => fs.mkdtempSync(path.join(os.tmpdir(), prefix)),
  joinPath: (...parts) => path.join(...parts),
  runAppCommandAttached: (appPath, appArgs, logLevel, label) =>
    runAppCommandAttached(appPath, appArgs, logLevel, label),
  launchAppCommandDetached: (appPath, appArgs, logLevel, label) =>
    launchAppCommandDetached(appPath, appArgs, logLevel, label),
  waitForStatsResponse: async (responsePath) => {
    const deadline = Date.now() + STATS_STARTUP_RESPONSE_TIMEOUT_MS;
    while (Date.now() < deadline) {
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

  try {
    const forwarded = ['--stats', '--stats-response-path', responsePath];
    if (args.statsBackground) {
      forwarded.push('--stats-background');
    }
    if (args.statsStop) {
      forwarded.push('--stats-stop');
    }
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
    if (args.statsBackground) {
      resolvedDeps.launchAppCommandDetached(appPath, forwarded, args.logLevel, 'stats');
      const startupResult = await resolvedDeps.waitForStatsResponse(responsePath);
      if (!startupResult.ok) {
        throw new Error(startupResult.error || 'Stats dashboard failed to start.');
      }
      return true;
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

    if (!args.statsCleanup && !args.statsStop) {
      const startupResult = await Promise.race([
        resolvedDeps
          .waitForStatsResponse(responsePath)
          .then((response) => ({ kind: 'response' as const, response })),
        attachedExitPromise.then((status) => ({ kind: 'exit' as const, status })),
      ]);
      if (startupResult.kind === 'exit') {
        if (startupResult.status !== 0) {
          throw new Error(
            `Stats app exited before startup response (status ${startupResult.status}).`,
          );
        }
        const response = await resolvedDeps.waitForStatsResponse(responsePath);
        if (!response.ok) {
          throw new Error(response.error || 'Stats dashboard failed to start.');
        }
        return true;
      }
      if (!startupResult.response.ok) {
        throw new Error(startupResult.response.error || 'Stats dashboard failed to start.');
      }
      await attachedExitPromise;
      return true;
    }
    const attachedExitPromiseCleanup = attachedExitPromise;

    const startupResult = await Promise.race([
      resolvedDeps
        .waitForStatsResponse(responsePath)
        .then((response) => ({ kind: 'response' as const, response })),
      attachedExitPromiseCleanup.then((status) => ({ kind: 'exit' as const, status })),
    ]);
    if (startupResult.kind === 'exit') {
      if (startupResult.status !== 0) {
        throw new Error(
          `Stats app exited before startup response (status ${startupResult.status}).`,
        );
      }
      const response = await resolvedDeps.waitForStatsResponse(responsePath);
      if (!response.ok) {
        throw new Error(response.error || 'Stats dashboard failed to start.');
      }
      return true;
    }
    if (!startupResult.response.ok) {
      throw new Error(startupResult.response.error || 'Stats dashboard failed to start.');
    }
    const exitStatus = await attachedExitPromiseCleanup;
    if (exitStatus !== 0) {
      throw new Error(`Stats app exited with status ${exitStatus}.`);
    }
    return true;
  } finally {
    resolvedDeps.removeDir(tempDir);
  }
}
