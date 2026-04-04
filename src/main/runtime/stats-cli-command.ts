import fs from 'node:fs';
import path from 'node:path';
import type { CliArgs, CliCommandSource } from '../../cli/args';
import type {
  LifetimeRebuildSummary,
  VocabularyCleanupSummary,
} from '../../core/services/immersion-tracker/types';

type StatsCliConfig = {
  immersionTracking?: {
    enabled?: boolean;
  };
  stats: {
    serverPort: number;
    autoOpenBrowser?: boolean;
  };
};

export type StatsCliCommandResponse = {
  ok: boolean;
  url?: string;
  error?: string;
};

type BackgroundStatsStartResult = {
  url: string;
  runningInCurrentProcess: boolean;
};

type BackgroundStatsStopResult = {
  ok: boolean;
  stale: boolean;
};

function formatLoggedNumber(value: number): string {
  return Number.isFinite(value) ? value.toString() : String(value);
}

export function writeStatsCliCommandResponse(
  responsePath: string,
  payload: StatsCliCommandResponse,
): void {
  fs.mkdirSync(path.dirname(responsePath), { recursive: true });
  fs.writeFileSync(responsePath, JSON.stringify(payload, null, 2), 'utf8');
}

export function createRunStatsCliCommandHandler(deps: {
  getResolvedConfig: () => StatsCliConfig;
  ensureImmersionTrackerStarted: () => void;
  ensureVocabularyCleanupTokenizerReady?: () => Promise<void> | void;
  getImmersionTracker: () => {
    cleanupVocabularyStats?: () => Promise<VocabularyCleanupSummary>;
    rebuildLifetimeSummaries?: () => Promise<LifetimeRebuildSummary>;
  } | null;
  ensureStatsServerStarted: () => string;
  ensureBackgroundStatsServerStarted: () => BackgroundStatsStartResult;
  stopBackgroundStatsServer: () => Promise<BackgroundStatsStopResult> | BackgroundStatsStopResult;
  openExternal: (url: string) => Promise<unknown>;
  writeResponse: (responsePath: string, payload: StatsCliCommandResponse) => void;
  exitAppWithCode: (code: number) => void;
  logInfo: (message: string) => void;
  logWarn: (message: string, error: unknown) => void;
  logError: (message: string, error: unknown) => void;
}) {
  const writeResponseSafe = (
    responsePath: string | undefined,
    payload: StatsCliCommandResponse,
  ): void => {
    if (!responsePath) return;
    try {
      deps.writeResponse(responsePath, payload);
    } catch (error) {
      deps.logWarn(`Failed to write stats response: ${responsePath}`, error);
    }
  };

  return async (
    args: Pick<
      CliArgs,
      | 'statsResponsePath'
      | 'statsBackground'
      | 'statsStop'
      | 'statsCleanup'
      | 'statsCleanupVocab'
      | 'statsCleanupLifetime'
    >,
    source: CliCommandSource,
  ): Promise<void> => {
    try {
      if (args.statsStop) {
        const result = await deps.stopBackgroundStatsServer();
        deps.logInfo(
          result.stale
            ? 'Background stats server is not running; cleaned stale state.'
            : 'Background stats server stopped.',
        );
        writeResponseSafe(args.statsResponsePath, { ok: true });
        if (source === 'initial') {
          deps.exitAppWithCode(0);
        }
        return;
      }

      const config = deps.getResolvedConfig();
      if (config.immersionTracking?.enabled === false) {
        throw new Error('Immersion tracking is disabled in config.');
      }

      if (args.statsBackground) {
        const result = deps.ensureBackgroundStatsServerStarted();
        deps.logInfo(`Stats dashboard available at ${result.url}`);
        writeResponseSafe(args.statsResponsePath, { ok: true, url: result.url });
        if (!result.runningInCurrentProcess && source === 'initial') {
          deps.exitAppWithCode(0);
        }
        return;
      }

      deps.ensureImmersionTrackerStarted();
      const tracker = deps.getImmersionTracker();
      if (!tracker) {
        throw new Error('Immersion tracker failed to initialize.');
      }

      if (args.statsCleanup) {
        const cleanupModes = [
          args.statsCleanupVocab ? 'vocab' : null,
          args.statsCleanupLifetime ? 'lifetime' : null,
        ].filter(Boolean);
        if (cleanupModes.length !== 1) {
          throw new Error('Choose exactly one stats cleanup mode.');
        }

        if (args.statsCleanupVocab) {
          await deps.ensureVocabularyCleanupTokenizerReady?.();
        }
        if (args.statsCleanupVocab && tracker.cleanupVocabularyStats) {
          const result = await tracker.cleanupVocabularyStats();
          deps.logInfo(
            `Stats vocabulary cleanup complete: scanned=${result.scanned} kept=${result.kept} deleted=${result.deleted} repaired=${result.repaired}`,
          );
          writeResponseSafe(args.statsResponsePath, { ok: true });
          return;
        }
        if (!args.statsCleanupLifetime || !tracker.rebuildLifetimeSummaries) {
          throw new Error('Stats cleanup mode is not available.');
        }
        const result = await tracker.rebuildLifetimeSummaries();
        deps.logInfo(
          `Stats lifetime rebuild complete: appliedSessions=${formatLoggedNumber(result.appliedSessions)} rebuiltAtMs=${formatLoggedNumber(result.rebuiltAtMs)}`,
        );
        writeResponseSafe(args.statsResponsePath, { ok: true });
        return;
      }

      const url = deps.ensureStatsServerStarted();
      if (config.stats.autoOpenBrowser !== false) {
        await deps.openExternal(url);
      }
      deps.logInfo(`Stats dashboard available at ${url}`);
      writeResponseSafe(args.statsResponsePath, { ok: true, url });
    } catch (error) {
      deps.logError('Stats command failed', error);
      const message = error instanceof Error ? error.message : String(error);
      writeResponseSafe(args.statsResponsePath, { ok: false, error: message });
      if (source === 'initial') {
        deps.exitAppWithCode(1);
      }
    }
  };
}
