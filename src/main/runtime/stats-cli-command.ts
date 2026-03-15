import fs from 'node:fs';
import path from 'node:path';
import type { CliArgs, CliCommandSource } from '../../cli/args';
import type { VocabularyCleanupSummary } from '../../core/services/immersion-tracker/types';

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
  getImmersionTracker: () => { cleanupVocabularyStats?: () => Promise<VocabularyCleanupSummary> } | null;
  ensureStatsServerStarted: () => string;
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
    args: Pick<CliArgs, 'statsResponsePath' | 'statsCleanup' | 'statsCleanupVocab'>,
    source: CliCommandSource,
  ): Promise<void> => {
    try {
      const config = deps.getResolvedConfig();
      if (config.immersionTracking?.enabled === false) {
        throw new Error('Immersion tracking is disabled in config.');
      }

      deps.ensureImmersionTrackerStarted();
      const tracker = deps.getImmersionTracker();
      if (!tracker) {
        throw new Error('Immersion tracker failed to initialize.');
      }

      if (args.statsCleanup) {
        await deps.ensureVocabularyCleanupTokenizerReady?.();
        if (!args.statsCleanupVocab || !tracker.cleanupVocabularyStats) {
          throw new Error('Stats cleanup mode is not available.');
        }
        const result = await tracker.cleanupVocabularyStats();
        deps.logInfo(
          `Stats vocabulary cleanup complete: scanned=${result.scanned} kept=${result.kept} deleted=${result.deleted} repaired=${result.repaired}`,
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
