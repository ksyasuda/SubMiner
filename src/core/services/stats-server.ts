import { Hono } from 'hono';
import http, { type IncomingMessage, type ServerResponse } from 'node:http';
import { Readable } from 'node:stream';
import type { AnkiConnectConfig } from '../../types.js';
import type { AnilistRateLimiter } from './anilist/rate-limiter.js';
import type { TmdbClient } from './tmdb/tmdb-client.js';
import type { ImmersionTrackerService } from './immersion-tracker-service.js';
import type { RetimedSecondarySubtitleInput } from './secondary-subtitle-sidecar.js';
import type { StatsServerMediaGenerator } from './stats-server/mining-support.js';
import { enforceStatsRequestSafety } from './stats-server/request-safety.js';
import {
  registerStatsAnalyticsRoutes,
  registerStatsIntegrationRoutes,
  registerStatsLibraryRoutes,
  registerStatsMiningRoutes,
  registerStatsStaticRoutes,
} from './stats-server/routes.js';

export type { StatsMiningTimingEvent } from './stats-server/mining-support.js';
import type { StatsMiningTimingEvent } from './stats-server/mining-support.js';

function toFetchHeaders(headers: IncomingMessage['headers']): Headers {
  const fetchHeaders = new Headers();
  for (const [name, value] of Object.entries(headers)) {
    if (value === undefined) continue;
    if (Array.isArray(value)) {
      for (const entry of value) fetchHeaders.append(name, entry);
      continue;
    }
    fetchHeaders.set(name, value);
  }
  return fetchHeaders;
}

function toFetchRequest(req: IncomingMessage): Request {
  const method = req.method ?? 'GET';
  const url = new URL(req.url ?? '/', `http://${req.headers.host ?? '127.0.0.1'}`);
  const init: RequestInit & { duplex?: 'half' } = {
    method,
    headers: toFetchHeaders(req.headers),
  };
  const hasBody =
    req.headers['transfer-encoding'] !== undefined ||
    Number(req.headers['content-length'] ?? 0) > 0;
  if (method !== 'GET' && method !== 'HEAD' && hasBody) {
    init.body = Readable.toWeb(req) as BodyInit;
    init.duplex = 'half';
  }
  return new Request(url, init);
}

async function writeFetchResponse(res: ServerResponse, response: Response): Promise<void> {
  res.statusCode = response.status;
  response.headers.forEach((value, key) => res.setHeader(key, value));
  res.end(Buffer.from(await response.arrayBuffer()));
}

export interface StatsServer {
  close: () => Promise<void>;
}

const SHUTDOWN_GRACE_MS = 1_000;

type BunServe = (options: {
  fetch: (typeof Hono.prototype)['fetch'];
  port: number;
  hostname: string;
}) => {
  stop: () => Promise<void> | void;
};

export function startNodeHttpServer(
  app: Hono,
  config: StatsServerConfig,
  createServer: (listener: http.RequestListener) => http.Server = http.createServer,
): Promise<StatsServer> {
  const server = createServer((req, res) => {
    void (async () => {
      try {
        await writeFetchResponse(res, await app.fetch(toFetchRequest(req)));
      } catch {
        res.statusCode = 500;
        res.end('Internal Server Error');
      }
    })();
  });
  return new Promise((resolve, reject) => {
    const handleStartupError = (error: Error): void => {
      server.removeListener('listening', handleListening);
      reject(error);
    };
    const handleListening = (): void => {
      server.removeListener('error', handleStartupError);
      let closePromise: Promise<void> | null = null;
      resolve({
        close: () => {
          closePromise ??= new Promise<void>((closeResolve, closeReject) => {
            const forceClose = setTimeout(() => server.closeAllConnections(), SHUTDOWN_GRACE_MS);
            server.close((error) => {
              clearTimeout(forceClose);
              if (error) closeReject(error);
              else closeResolve();
            });
          });
          return closePromise;
        },
      });
    };

    server.once('error', handleStartupError);
    server.once('listening', handleListening);
    server.listen(config.port, '127.0.0.1');
  });
}

export interface StatsServerConfig {
  port: number;
  staticDir: string; // Path to stats/dist/
  tracker: ImmersionTrackerService;
  knownWordCachePath?: string;
  mpvSocketPath?: string;
  ankiConnectConfig?: AnkiConnectConfig;
  getAnkiConnectConfig?: () => AnkiConnectConfig | undefined;
  getYomitanAnkiDeckName?: () => Promise<string | null | undefined> | string | null | undefined;
  secondarySubtitleLanguages?: string[];
  getSecondarySubtitleLanguages?: () => string[] | undefined;
  statsMiningAlassPath?: string;
  getStatsMiningAlassPath?: () => string | null | undefined;
  resolveRetimedSecondarySubtitleText?: (
    input: RetimedSecondarySubtitleInput,
  ) => Promise<string> | string;
  anilistRateLimiter?: AnilistRateLimiter;
  tmdbClient?: TmdbClient;
  addYomitanNote?: (word: string) => Promise<number | null>;
  resolveAnkiNoteId?: (noteId: number) => number;
  resolveSentenceSearchHeadwords?: (term: string) => Promise<string[]> | string[];
}

export function createStatsApp(
  tracker: ImmersionTrackerService,
  options?: {
    staticDir?: string;
    knownWordCachePath?: string;
    mpvSocketPath?: string;
    ankiConnectConfig?: AnkiConnectConfig;
    getAnkiConnectConfig?: () => AnkiConnectConfig | undefined;
    getYomitanAnkiDeckName?: () => Promise<string | null | undefined> | string | null | undefined;
    secondarySubtitleLanguages?: string[];
    getSecondarySubtitleLanguages?: () => string[] | undefined;
    statsMiningAlassPath?: string;
    getStatsMiningAlassPath?: () => string | null | undefined;
    resolveRetimedSecondarySubtitleText?: (
      input: RetimedSecondarySubtitleInput,
    ) => Promise<string> | string;
    anilistRateLimiter?: AnilistRateLimiter;
    tmdbClient?: TmdbClient;
    addYomitanNote?: (word: string) => Promise<number | null>;
    resolveAnkiNoteId?: (noteId: number) => number;
    resolveSentenceSearchHeadwords?: (term: string) => Promise<string[]> | string[];
    createMediaGenerator?: () => StatsServerMediaGenerator;
    onMiningTiming?: (event: StatsMiningTimingEvent) => void;
    nowMs?: () => number;
  },
) {
  const app = new Hono();
  app.use('*', enforceStatsRequestSafety);
  registerStatsAnalyticsRoutes(app, tracker, options);
  registerStatsLibraryRoutes(app, tracker, options);
  registerStatsIntegrationRoutes(app, tracker, options);
  registerStatsMiningRoutes(app, options);
  registerStatsStaticRoutes(app, options?.staticDir);
  return app;
}

export async function startStatsServerWithRuntime(
  config: StatsServerConfig,
  runtime: { bunServe: BunServe | null },
): Promise<StatsServer> {
  const app = createStatsApp(config.tracker, {
    staticDir: config.staticDir,
    knownWordCachePath: config.knownWordCachePath,
    mpvSocketPath: config.mpvSocketPath,
    ankiConnectConfig: config.ankiConnectConfig,
    getAnkiConnectConfig: config.getAnkiConnectConfig,
    getYomitanAnkiDeckName: config.getYomitanAnkiDeckName,
    secondarySubtitleLanguages: config.secondarySubtitleLanguages,
    getSecondarySubtitleLanguages: config.getSecondarySubtitleLanguages,
    statsMiningAlassPath: config.statsMiningAlassPath,
    getStatsMiningAlassPath: config.getStatsMiningAlassPath,
    resolveRetimedSecondarySubtitleText: config.resolveRetimedSecondarySubtitleText,
    anilistRateLimiter: config.anilistRateLimiter,
    tmdbClient: config.tmdbClient,
    addYomitanNote: config.addYomitanNote,
    resolveAnkiNoteId: config.resolveAnkiNoteId,
    resolveSentenceSearchHeadwords: config.resolveSentenceSearchHeadwords,
  });

  if (runtime.bunServe) {
    const server = runtime.bunServe({
      fetch: app.fetch,
      port: config.port,
      hostname: '127.0.0.1',
    });
    let closePromise: Promise<void> | null = null;
    return Promise.resolve({
      close: () => {
        closePromise ??= Promise.resolve().then(() => server.stop());
        return closePromise;
      },
    });
  }
  return startNodeHttpServer(app, config);
}

export function startStatsServer(config: StatsServerConfig): Promise<StatsServer> {
  const bunRuntime = globalThis as typeof globalThis & {
    Bun?: { serve?: BunServe };
  };
  return startStatsServerWithRuntime(config, { bunServe: bunRuntime.Bun?.serve ?? null });
}
