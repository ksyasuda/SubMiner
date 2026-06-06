import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { ConfigService } from './config/service';
import { createLogger, setLogLevel } from './logger';
import { ImmersionTrackerService } from './core/services/immersion-tracker-service';
import { createCoverArtFetcher } from './core/services/anilist/cover-art-fetcher';
import { createAnilistRateLimiter } from './core/services/anilist/rate-limiter';
import { startStatsServer } from './core/services/stats-server';
import {
  removeBackgroundStatsServerState,
  writeBackgroundStatsServerState,
} from './main/runtime/stats-daemon';
import { writeStatsCliCommandResponse } from './main/runtime/stats-cli-command';
import {
  createInvokeStatsWordHelperHandler,
  createReadStatsYomitanDeckNameHandler,
  type StatsWordHelperSpawnOptions,
  type StatsWordHelperResponse,
} from './stats-word-helper-client';

const logger = createLogger('stats-daemon');
const STATS_WORD_HELPER_RESPONSE_TIMEOUT_MS = 20_000;

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

async function waitForWordHelperResponse(responsePath: string): Promise<StatsWordHelperResponse> {
  const deadline = Date.now() + STATS_WORD_HELPER_RESPONSE_TIMEOUT_MS;
  while (Date.now() < deadline) {
    try {
      if (fs.existsSync(responsePath)) {
        return JSON.parse(fs.readFileSync(responsePath, 'utf8')) as StatsWordHelperResponse;
      }
    } catch {
      // retry until timeout
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }

  return {
    ok: false,
    error: 'Timed out waiting for stats word helper response.',
  };
}

const statsWordHelperDeps = {
  createTempDir: (prefix: string) => fs.mkdtempSync(path.join(os.tmpdir(), prefix)),
  joinPath: (...parts: string[]) => path.join(...parts),
  spawnHelper: async (options: StatsWordHelperSpawnOptions) => {
    const childArgs = [
      options.scriptPath,
      '--stats-word-helper-response-path',
      options.responsePath,
      '--stats-word-helper-user-data-path',
      options.userDataPath,
    ];
    if (options.mode === 'deck-name') {
      childArgs.push('--stats-word-helper-read-deck');
    } else {
      childArgs.push('--stats-word-helper-word', options.word ?? '');
    }
    const logLevel = readFlagValue(process.argv, '--log-level');
    if (logLevel) {
      childArgs.push('--log-level', logLevel);
    }
    const child = spawn(process.execPath, childArgs, {
      stdio: 'ignore',
      env: {
        ...process.env,
        ELECTRON_RUN_AS_NODE: undefined,
      },
    });
    return await new Promise<number>((resolve) => {
      child.once('exit', (code) => resolve(code ?? 1));
      child.once('error', () => resolve(1));
    });
  },
  waitForResponse: waitForWordHelperResponse,
  removeDir: (targetPath: string) => {
    fs.rmSync(targetPath, { recursive: true, force: true });
  },
};

const invokeStatsWordHelper = createInvokeStatsWordHelperHandler(statsWordHelperDeps);
const readStatsYomitanDeckName = createReadStatsYomitanDeckNameHandler(statsWordHelperDeps);

const userDataPath = readFlagValue(process.argv, '--stats-user-data-path')?.trim();
const responsePath = readFlagValue(process.argv, '--stats-response-path')?.trim();
const logLevel = readFlagValue(process.argv, '--log-level');

if (logLevel) {
  setLogLevel(logLevel, 'cli');
}

if (!userDataPath) {
  if (responsePath) {
    writeStatsCliCommandResponse(responsePath, {
      ok: false,
      error: 'Missing --stats-user-data-path for stats daemon runner.',
    });
  }
  process.exit(1);
}

const daemonUserDataPath = userDataPath;

const statePath = path.join(userDataPath, 'stats-daemon.json');
const knownWordCachePath = path.join(userDataPath, 'known-words-cache.json');
const statsDistPath = path.join(__dirname, '..', 'stats', 'dist');
const wordHelperScriptPath = path.join(__dirname, 'stats-word-helper.js');

let tracker: ImmersionTrackerService | null = null;
let statsServer: ReturnType<typeof startStatsServer> | null = null;

function writeFailureResponse(message: string): void {
  if (!responsePath) return;
  writeStatsCliCommandResponse(responsePath, { ok: false, error: message });
}

function clearOwnedState(): void {
  const rawState = (() => {
    try {
      return JSON.parse(fs.readFileSync(statePath, 'utf8')) as { pid?: number };
    } catch {
      return null;
    }
  })();
  if (rawState?.pid === process.pid) {
    removeBackgroundStatsServerState(statePath);
  }
}

function shutdown(code = 0): void {
  try {
    statsServer?.close();
  } catch {
    // ignore
  }
  statsServer = null;
  try {
    tracker?.destroy();
  } catch {
    // ignore
  }
  tracker = null;
  clearOwnedState();
  process.exit(code);
}

process.on('SIGINT', () => shutdown(0));
process.on('SIGTERM', () => shutdown(0));

async function main(): Promise<void> {
  try {
    const configService = new ConfigService(daemonUserDataPath);
    const config = configService.getConfig();
    if (config.immersionTracking?.enabled === false) {
      throw new Error('Immersion tracking is disabled in config.');
    }

    const configuredDbPath = config.immersionTracking?.dbPath?.trim() || '';
    tracker = new ImmersionTrackerService({
      dbPath: configuredDbPath || path.join(daemonUserDataPath, 'immersion.sqlite'),
      policy: {
        batchSize: config.immersionTracking.batchSize,
        flushIntervalMs: config.immersionTracking.flushIntervalMs,
        queueCap: config.immersionTracking.queueCap,
        payloadCapBytes: config.immersionTracking.payloadCapBytes,
        maintenanceIntervalMs: config.immersionTracking.maintenanceIntervalMs,
        retention: {
          eventsDays: config.immersionTracking.retention.eventsDays,
          telemetryDays: config.immersionTracking.retention.telemetryDays,
          sessionsDays: config.immersionTracking.retention.sessionsDays,
          dailyRollupsDays: config.immersionTracking.retention.dailyRollupsDays,
          monthlyRollupsDays: config.immersionTracking.retention.monthlyRollupsDays,
          vacuumIntervalDays: config.immersionTracking.retention.vacuumIntervalDays,
        },
      },
    });
    tracker.setCoverArtFetcher(
      createCoverArtFetcher(createAnilistRateLimiter(), createLogger('stats-daemon:cover-art')),
    );

    statsServer = startStatsServer({
      port: config.stats.serverPort,
      staticDir: statsDistPath,
      tracker,
      knownWordCachePath,
      getAnkiConnectConfig: () => configService.reloadConfig().ankiConnect,
      getYomitanAnkiDeckName: async () =>
        await readStatsYomitanDeckName({
          helperScriptPath: wordHelperScriptPath,
          userDataPath: daemonUserDataPath,
        }),
      getSecondarySubtitleLanguages: () =>
        configService.reloadConfig().secondarySub.secondarySubLanguages,
      addYomitanNote: async (word: string) =>
        await invokeStatsWordHelper({
          helperScriptPath: wordHelperScriptPath,
          userDataPath: daemonUserDataPath,
          word,
        }),
    });

    writeBackgroundStatsServerState(statePath, {
      pid: process.pid,
      port: config.stats.serverPort,
      startedAtMs: Date.now(),
    });

    if (responsePath) {
      writeStatsCliCommandResponse(responsePath, {
        ok: true,
        url: `http://127.0.0.1:${config.stats.serverPort}`,
      });
    }
    logger.info(`Background stats daemon listening on http://127.0.0.1:${config.stats.serverPort}`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('Failed to start stats daemon', message);
    writeFailureResponse(message);
    shutdown(1);
  }
}

void main();
