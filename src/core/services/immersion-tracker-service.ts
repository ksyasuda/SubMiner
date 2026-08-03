import path from 'node:path';
import * as fs from 'node:fs';
import { createLogger } from '../../logger';
import { MediaGenerator } from '../../media-generator';
import type { CoverArtFetcher } from './anilist/cover-art-fetcher';
import { getLocalVideoMetadata, guessAnimeVideoMetadata } from './immersion-tracker/metadata';
import {
  pruneRawRetention,
  pruneRollupRetention,
  runOptimizeMaintenance,
  runRollupMaintenance,
} from './immersion-tracker/maintenance';
import { Database, type DatabaseSync } from './immersion-tracker/sqlite';
import { finalizeSessionRecord, startSessionRecord } from './immersion-tracker/session';
import {
  applyPragmas,
  createTrackerPreparedStatements,
  ensureSchema,
  findManualDirectoryAnimeAssignment,
  getManualAnimeAssignment,
  executeQueuedWrite,
  getOrCreateAnimeRecord,
  getOrCreateVideoRecord,
  linkVideoToAnimeRecord,
  linkYoutubeVideoToAnimeRecord,
  type TrackerPreparedStatements,
  updateVideoMetadataRecord,
  updateVideoTitleRecord,
  upsertYoutubeVideoMetadata,
} from './immersion-tracker/storage';
import {
  applySessionLifetimeSummary,
  reconcileStaleActiveSessions,
  rebuildLifetimeSummaries as rebuildLifetimeSummaryTables,
  rebuildLifetimeSummariesInTransaction,
  recomputeLifetimeAnimeFromMedia,
  recomputeLifetimeGlobalFromSummaries,
  repairLifetimeSummariesFromMedia,
  shouldBackfillLifetimeSummaries,
} from './immersion-tracker/lifetime';
import {
  getAllDistinctHeadwords,
  getAnimeDistinctHeadwords,
  getDailyRollups,
  getMediaDistinctHeadwords,
  getMonthlyRollups,
  getQueryHints,
  getSessionSummaries,
  getSessionTimeline,
  getSessionWordsByLine,
} from './immersion-tracker/query-sessions';
import { getTrendsDashboard } from './immersion-tracker/query-trends';
import {
  getKanjiAnimeAppearances,
  getKanjiDetail,
  getKanjiOccurrences,
  getKanjiStats,
  getKanjiWords,
  getSessionEvents,
  getSimilarWords,
  getStatsExcludedWords,
  getVocabularyStats,
  replaceStatsExcludedWords,
  searchSubtitleSentences,
  getWordAnimeAppearances,
  getWordDetail,
  getWordOccurrences,
} from './immersion-tracker/query-lexical';
import {
  getAnimeAnilistEntries,
  getAnimeCoverArt,
  getAnimeDailyRollups,
  getAnimeDetail,
  getAnimeEpisodes,
  getAnimeLibrary,
  getAnimeWords,
  getCoverArt,
  getEpisodeCardEvents,
  getEpisodeSessions,
  getEpisodeWords,
  getEpisodesPerDay,
  getMediaDailyRollups,
  getMediaDetail,
  getMediaLibrary,
  getMediaSessions,
  getNewAnimePerDay,
  getStreakCalendar,
  getWatchTimePerAnime,
} from './immersion-tracker/query-library';
import {
  cleanupVocabularyStats,
  getVideoDurationMs,
  markVideoWatched,
  upsertCoverArt,
} from './immersion-tracker/query-maintenance';
import {
  DeleteMaintenanceWorkerRuntime,
  type RunDeleteMaintenanceTask,
} from './immersion-tracker/delete-maintenance-worker-runtime';
import { DeleteMaintenanceScheduler } from './immersion-tracker/delete-maintenance-scheduler';
import {
  cleanupDuplicateSubtitleLines,
  type DuplicateSubtitleLineCleanupOptions,
  type DuplicateSubtitleLineCleanupSummary,
} from './immersion-tracker/duplicate-line-cleanup';
import {
  getVideoIdByVideoKey,
  getWatchStateByVideoKeys,
} from './immersion-tracker/query-watch-state';
import { repairJellyfinStreamVideoLinks } from './immersion-tracker/jellyfin-link-repair';
import {
  dismissAnimeMergeRecommendation,
  getAnimeMergeRecommendations,
  repairLegacySeasonlessAnimeRows,
  resolveAnimeAnilistConflict,
  type AnimeMergeRecommendation,
} from './immersion-tracker/anime-season-repair';
import {
  mergeAnimeRecords,
  moveVideoToAnime as moveVideoToAnimeQuery,
  type AnimeMergeSummary,
  type VideoMoveSummary,
} from './immersion-tracker/anime-merge';
import {
  buildVideoKey,
  deriveCanonicalTitle,
  isKanji,
  isRemoteSource,
  normalizeMediaPath,
  normalizeText,
  resolveBoundedInt,
  sanitizePayload,
  secToMs,
} from './immersion-tracker/reducer';
import { DEFAULT_MIN_WATCH_RATIO } from '../../shared/watch-threshold';
import { enqueueWrite } from './immersion-tracker/queue';
import { nowMs } from './immersion-tracker/time';
import {
  DEFAULT_BATCH_SIZE,
  DEFAULT_DAILY_ROLLUP_RETENTION_MS,
  DEFAULT_EVENTS_RETENTION_MS,
  DEFAULT_FLUSH_INTERVAL_MS,
  DEFAULT_MAINTENANCE_INTERVAL_MS,
  DEFAULT_MAX_PAYLOAD_BYTES,
  DEFAULT_MONTHLY_ROLLUP_RETENTION_MS,
  DEFAULT_QUEUE_CAP,
  DEFAULT_SESSIONS_RETENTION_MS,
  DEFAULT_TELEMETRY_RETENTION_MS,
  DEFAULT_VACUUM_INTERVAL_MS,
  EVENT_CARD_MINED,
  EVENT_LOOKUP,
  EVENT_MEDIA_BUFFER,
  EVENT_PAUSE_END,
  EVENT_PAUSE_START,
  EVENT_SEEK_BACKWARD,
  EVENT_SEEK_FORWARD,
  EVENT_SUBTITLE_LINE,
  EVENT_YOMITAN_LOOKUP,
  SOURCE_TYPE_LOCAL,
  SOURCE_TYPE_REMOTE,
  type ImmersionSessionRollupRow,
  type EpisodeCardEventRow,
  type EpisodesPerDayRow,
  type ImmersionTrackerOptions,
  type KanjiAnimeAppearanceRow,
  type KanjiDetailRow,
  type KanjiOccurrenceRow,
  type KanjiStatsRow,
  type KanjiWordRow,
  type LifetimeRebuildSummary,
  type LegacyVocabularyPosResolution,
  type LegacyVocabularyPosRow,
  type AnimeAnilistEntryRow,
  type AnimeDetailRow,
  type AnimeEpisodeRow,
  type AnimeLibraryRow,
  type AnimeWordRow,
  type MediaArtRow,
  type MediaDetailRow,
  type MediaLibraryRow,
  type NewAnimePerDayRow,
  type QueuedWrite,
  type SentenceSearchOptions,
  type SentenceSearchResultRow,
  type SessionEventRow,
  type SessionState,
  type SessionSummaryQueryRow,
  type SessionTimelineRow,
  type SimilarWordRow,
  type StatsExcludedWordRow,
  type StreakCalendarRow,
  type VocabularyCleanupSummary,
  type WatchTimePerAnimeRow,
  type WordAnimeAppearanceRow,
  type WordDetailRow,
  type WordOccurrenceRow,
  type VocabularyStatsRow,
  type CountedWordOccurrence,
} from './immersion-tracker/types';
import type { MergedToken } from '../../types';
import { shouldExcludeTokenFromVocabularyPersistence } from './tokenizer/annotation-stage';
import { deriveStoredPartOfSpeech } from './tokenizer/part-of-speech';
import { probeYoutubeVideoMetadata } from './youtube/metadata-probe';

const YOUTUBE_COVER_RETRY_MS = 5 * 60 * 1000;
const YOUTUBE_SCREENSHOT_MAX_SECONDS = 120;
const YOUTUBE_OEMBED_ENDPOINT = 'https://www.youtube.com/oembed';
const YOUTUBE_ID_PATTERN = /^[A-Za-z0-9_-]{6,}$/;
const YOUTUBE_METADATA_REFRESH_MS = 24 * 60 * 60 * 1000;
const DELETE_MAINTENANCE_BATCH_WINDOW_MS = 10;

function isValidYouTubeVideoId(value: string | null): boolean {
  return Boolean(value && YOUTUBE_ID_PATTERN.test(value));
}

function extractYouTubeVideoId(mediaUrl: string): string | null {
  let parsed: URL;
  try {
    parsed = new URL(mediaUrl);
  } catch {
    return null;
  }

  const host = parsed.hostname.toLowerCase();
  if (
    host !== 'youtu.be' &&
    !host.endsWith('.youtu.be') &&
    !host.endsWith('youtube.com') &&
    !host.endsWith('youtube-nocookie.com')
  ) {
    return null;
  }

  if (host === 'youtu.be' || host.endsWith('.youtu.be')) {
    const pathId = parsed.pathname.split('/').filter(Boolean)[0];
    return isValidYouTubeVideoId(pathId ?? null) ? (pathId as string) : null;
  }

  const queryId = parsed.searchParams.get('v') ?? parsed.searchParams.get('vi') ?? null;
  if (isValidYouTubeVideoId(queryId)) {
    return queryId;
  }

  const pathParts = parsed.pathname.split('/').filter(Boolean);
  for (let i = 0; i < pathParts.length; i += 1) {
    const current = pathParts[i];
    const next = pathParts[i + 1];
    if (!current || !next) continue;
    if (
      current.toLowerCase() === 'shorts' ||
      current.toLowerCase() === 'embed' ||
      current.toLowerCase() === 'live' ||
      current.toLowerCase() === 'v'
    ) {
      const candidate = decodeURIComponent(next);
      if (isValidYouTubeVideoId(candidate)) {
        return candidate;
      }
    }
  }

  return null;
}

function buildYouTubeThumbnailUrls(videoId: string): string[] {
  return [
    `https://i.ytimg.com/vi/${videoId}/maxresdefault.jpg`,
    `https://i.ytimg.com/vi/${videoId}/hqdefault.jpg`,
    `https://i.ytimg.com/vi/${videoId}/sddefault.jpg`,
    `https://i.ytimg.com/vi/${videoId}/mqdefault.jpg`,
    `https://i.ytimg.com/vi/${videoId}/0.jpg`,
    `https://i.ytimg.com/vi/${videoId}/default.jpg`,
  ];
}

async function fetchYouTubeOEmbedThumbnail(mediaUrl: string): Promise<string | null> {
  try {
    const response = await fetch(
      `${YOUTUBE_OEMBED_ENDPOINT}?url=${encodeURIComponent(mediaUrl)}&format=json`,
    );
    if (!response.ok) {
      return null;
    }
    const payload = (await response.json()) as { thumbnail_url?: unknown };
    const candidate = typeof payload.thumbnail_url === 'string' ? payload.thumbnail_url.trim() : '';
    return candidate || null;
  } catch {
    return null;
  }
}

async function downloadImage(url: string): Promise<Buffer | null> {
  try {
    const response = await fetch(url);
    if (!response.ok) return null;
    const contentType = response.headers.get('content-type');
    if (contentType && !contentType.toLowerCase().startsWith('image/')) {
      return null;
    }
    return Buffer.from(await response.arrayBuffer());
  } catch {
    return null;
  }
}

export type {
  AnimeAnilistEntryRow,
  AnimeDetailRow,
  AnimeEpisodeRow,
  AnimeLibraryRow,
  AnimeWordRow,
  EpisodeCardEventRow,
  EpisodesPerDayRow,
  ImmersionSessionRollupRow,
  ImmersionTrackerOptions,
  ImmersionTrackerPolicy,
  KanjiAnimeAppearanceRow,
  KanjiDetailRow,
  KanjiOccurrenceRow,
  KanjiStatsRow,
  KanjiWordRow,
  MediaArtRow,
  MediaDetailRow,
  MediaLibraryRow,
  NewAnimePerDayRow,
  SessionEventRow,
  SessionSummaryQueryRow,
  SessionTimelineRow,
  SimilarWordRow,
  StatsExcludedWordRow,
  StreakCalendarRow,
  WatchTimePerAnimeRow,
  WordAnimeAppearanceRow,
  WordDetailRow,
  WordOccurrenceRow,
  VocabularyStatsRow,
} from './immersion-tracker/types';

export interface JellyfinPlaybackMetadataInput {
  mediaPath: string;
  displayTitle: string;
  itemTitle: string;
  seriesTitle: string | null;
  seasonNumber: number | null;
  episodeNumber: number | null;
  itemId: string;
}

/**
 * An episode the anime browser is about to stream.
 *
 * Reported up front for the same reason the Jellyfin variant is: the stream URL
 * mpv sees is a per-playback proxy address ending in `.m3u8`, so a session
 * started from it alone lands in a series named after the file extension.
 */
export interface StreamPlaybackMetadataInput {
  /** The URL handed to mpv; recorded as an alias of `statsPath`. */
  mediaPath: string;
  /** Stable per-episode identity. Survives the proxy port and token changing. */
  statsPath: string;
  displayTitle: string;
  seriesTitle: string;
  seasonNumber: number | null;
  episodeNumber: number | null;
}

/** What a caller needs to show "already watched" against a streamed episode. */
export interface StreamWatchState {
  /** Set once a session passed the completion threshold, or marked by hand. */
  watched: boolean;
  /** Start of the most recent session, or null when it was never played. */
  lastWatchedMs: number | null;
  sessionCount: number;
}

/**
 * Parser sources that are recorded before playback starts. A video carrying one
 * already has better metadata than filename guessing could produce, so the
 * guess is skipped rather than allowed to overwrite it.
 */
const PREPLAYBACK_PARSER_SOURCES = new Set(['jellyfin', 'anime-browser']);

function normalizeMetadataInt(value: number | null | undefined): number | null {
  return typeof value === 'number' && Number.isSafeInteger(value) ? value : null;
}

function buildJellyfinStatsMediaPath(mediaPath: string, itemId: string): string {
  const normalizedItemId = normalizeText(itemId);
  if (!normalizedItemId) {
    return mediaPath;
  }
  try {
    const parsed = new URL(mediaPath);
    return `jellyfin://${parsed.host}/item/${encodeURIComponent(normalizedItemId)}`;
  } catch {
    return `jellyfin://item/${encodeURIComponent(normalizedItemId)}`;
  }
}

const JELLYFIN_MEDIA_ALIAS_QUERY_KEYS = [
  'api_key',
  'StartTimeTicks',
  'AudioStreamIndex',
  'SubtitleStreamIndex',
];

function deleteSearchParamsCaseInsensitive(searchParams: URLSearchParams, names: string[]): void {
  const loweredNames = new Set(names.map((name) => name.toLowerCase()));
  for (const key of [...searchParams.keys()]) {
    if (loweredNames.has(key.toLowerCase())) {
      searchParams.delete(key);
    }
  }
}

function buildMediaPathAliasCandidates(mediaPath: string): string[] {
  const candidates = new Set<string>([mediaPath]);
  try {
    const parsed = new URL(mediaPath);
    deleteSearchParamsCaseInsensitive(parsed.searchParams, JELLYFIN_MEDIA_ALIAS_QUERY_KEYS);
    candidates.add(parsed.toString());
  } catch {
    // Non-URL fallback paths are already represented by the raw candidate.
  }
  return [...candidates];
}

export class ImmersionTrackerService {
  private readonly logger = createLogger('main:immersion-tracker');
  private readonly db: DatabaseSync;
  private readonly queue: QueuedWrite[] = [];
  private readonly queueCap: number;
  private readonly batchSize: number;
  private readonly flushIntervalMs: number;
  private readonly maintenanceIntervalMs: number;
  private readonly maxPayloadBytes: number;
  private readonly eventsRetentionMs: number;
  private readonly telemetryRetentionMs: number;
  private readonly sessionsRetentionMs: number;
  private readonly eventsRetentionDays: number | null;
  private readonly telemetryRetentionDays: number | null;
  private readonly sessionsRetentionDays: number | null;
  private readonly dailyRollupRetentionMs: number;
  private readonly monthlyRollupRetentionMs: number;
  private readonly vacuumIntervalMs: number;
  private readonly dbPath: string;
  private readonly writeLock = { locked: false };
  private readonly destroyDeleteMaintenanceRunner: () => void;
  private readonly deleteMaintenanceScheduler: DeleteMaintenanceScheduler;
  private flushTimer: ReturnType<typeof setTimeout> | null = null;
  private maintenanceTimer: ReturnType<typeof setInterval> | null = null;
  private flushScheduled = false;
  private droppedWriteCount = 0;
  private lastVacuumMs = 0;
  private isDestroyed = false;
  private sessionState: SessionState | null = null;
  private currentVideoKey = '';
  private currentMediaPathOrUrl = '';
  private readonly mediaGenerator = new MediaGenerator();
  private readonly preparedStatements: TrackerPreparedStatements;
  private coverArtFetcher: CoverArtFetcher | null = null;
  private readonly pendingCoverFetches = new Map<number, Promise<boolean>>();
  private readonly pendingYoutubeMetadataFetches = new Map<number, Promise<void>>();
  private readonly recordedSubtitleKeys = new Set<string>();
  private readonly pendingAnimeMetadataUpdates = new Map<number, Promise<void>>();
  private readonly mediaPathAliases = new Map<string, string>();
  private readonly resolveLegacyVocabularyPos:
    | ((row: LegacyVocabularyPosRow) => Promise<LegacyVocabularyPosResolution | null>)
    | undefined;

  constructor(
    options: ImmersionTrackerOptions,
    dependencies: {
      runDeleteMaintenanceTask?: RunDeleteMaintenanceTask;
      destroyDeleteMaintenanceRunner?: () => void;
    } = {},
  ) {
    this.dbPath = options.dbPath;
    this.resolveLegacyVocabularyPos = options.resolveLegacyVocabularyPos;
    let runDeleteMaintenanceTask: RunDeleteMaintenanceTask;
    if (dependencies.runDeleteMaintenanceTask) {
      runDeleteMaintenanceTask = dependencies.runDeleteMaintenanceTask;
      this.destroyDeleteMaintenanceRunner =
        dependencies.destroyDeleteMaintenanceRunner ?? (() => {});
    } else {
      const deleteMaintenanceRuntime = new DeleteMaintenanceWorkerRuntime();
      runDeleteMaintenanceTask = (dbPath, task) => deleteMaintenanceRuntime.run(dbPath, task);
      this.destroyDeleteMaintenanceRunner = () => deleteMaintenanceRuntime.destroy();
    }
    this.deleteMaintenanceScheduler = new DeleteMaintenanceScheduler({
      batchWindowMs: DELETE_MAINTENANCE_BATCH_WINDOW_MS,
      runTask: (task) => runDeleteMaintenanceTask(this.dbPath, task),
      onBusy: () => {
        this.requireWriteQueueDrained('delete maintenance');
        this.writeLock.locked = true;
      },
      onIdle: () => {
        this.writeLock.locked = false;
        if (!this.isDestroyed && this.queue.length > 0) this.scheduleFlush(0);
      },
    });
    const parentDir = path.dirname(this.dbPath);
    if (!fs.existsSync(parentDir)) {
      fs.mkdirSync(parentDir, { recursive: true });
    }

    const policy = options.policy ?? {};
    this.queueCap = resolveBoundedInt(policy.queueCap, DEFAULT_QUEUE_CAP, 100, 100_000);
    this.batchSize = resolveBoundedInt(policy.batchSize, DEFAULT_BATCH_SIZE, 1, 10_000);
    this.flushIntervalMs = resolveBoundedInt(
      policy.flushIntervalMs,
      DEFAULT_FLUSH_INTERVAL_MS,
      50,
      60_000,
    );
    this.maintenanceIntervalMs = resolveBoundedInt(
      policy.maintenanceIntervalMs,
      DEFAULT_MAINTENANCE_INTERVAL_MS,
      60_000,
      7 * 24 * 60 * 60 * 1000,
    );
    this.maxPayloadBytes = resolveBoundedInt(
      policy.payloadCapBytes,
      DEFAULT_MAX_PAYLOAD_BYTES,
      64,
      8192,
    );

    const retention = policy.retention ?? {};
    const daysToRetentionWindow = (
      value: number | undefined,
      fallbackDays: number,
      maxDays: number,
    ): { ms: number; days: number | null } => {
      const resolvedDays = resolveBoundedInt(value, fallbackDays, 0, maxDays);
      return {
        ms: resolvedDays === 0 ? Number.POSITIVE_INFINITY : resolvedDays * 86_400_000,
        days: resolvedDays === 0 ? null : resolvedDays,
      };
    };

    const eventsRetention = daysToRetentionWindow(retention.eventsDays, 7, 3650);
    const telemetryRetention = daysToRetentionWindow(retention.telemetryDays, 30, 3650);
    const sessionsRetention = daysToRetentionWindow(retention.sessionsDays, 30, 3650);
    this.eventsRetentionMs = eventsRetention.ms;
    this.eventsRetentionDays = eventsRetention.days;
    this.telemetryRetentionMs = telemetryRetention.ms;
    this.telemetryRetentionDays = telemetryRetention.days;
    this.sessionsRetentionMs = sessionsRetention.ms;
    this.sessionsRetentionDays = sessionsRetention.days;
    this.dailyRollupRetentionMs = daysToRetentionWindow(retention.dailyRollupsDays, 365, 36500).ms;
    this.monthlyRollupRetentionMs = daysToRetentionWindow(
      retention.monthlyRollupsDays,
      5 * 365,
      36500,
    ).ms;
    this.vacuumIntervalMs = daysToRetentionWindow(retention.vacuumIntervalDays, 7, 3650).ms;
    this.db = new Database(this.dbPath);
    applyPragmas(this.db);
    ensureSchema(this.db);
    const reconciledSessions = reconcileStaleActiveSessions(this.db);
    if (reconciledSessions > 0) {
      this.logger.info(
        `Recovered stale active sessions on startup: reconciledSessions=${reconciledSessions}`,
      );
    }
    const jellyfinRepair = repairJellyfinStreamVideoLinks(this.db);
    if (jellyfinRepair.repaired > 0) {
      this.logger.info(
        `Repaired Jellyfin stats links on startup: scanned=${jellyfinRepair.scanned} repaired=${jellyfinRepair.repaired}`,
      );
    }
    const seasonRepair = repairLegacySeasonlessAnimeRows(this.db);
    if (seasonRepair.movedVideos > 0 || seasonRepair.deletedAnimeRows > 0) {
      this.logger.info(
        `Repaired season-scoped stats links on startup: scanned=${seasonRepair.scanned} movedVideos=${seasonRepair.movedVideos} deletedAnimeRows=${seasonRepair.deletedAnimeRows}`,
      );
      repairLifetimeSummariesFromMedia(this.db);
    }
    if (shouldBackfillLifetimeSummaries(this.db)) {
      const result = rebuildLifetimeSummaryTables(this.db);
      if (result.appliedSessions > 0) {
        this.logger.info(
          `Backfilled lifetime summaries from retained sessions: appliedSessions=${result.appliedSessions}`,
        );
      }
    }
    this.preparedStatements = createTrackerPreparedStatements(this.db);
    this.scheduleMaintenance();
    this.scheduleFlush();
  }

  destroy(): void {
    if (this.isDestroyed) return;
    if (this.flushTimer) {
      clearTimeout(this.flushTimer);
      this.flushTimer = null;
    }
    if (this.maintenanceTimer) {
      clearInterval(this.maintenanceTimer);
      this.maintenanceTimer = null;
    }
    this.finalizeActiveSession();
    this.isDestroyed = true;
    this.deleteMaintenanceScheduler.destroy();
    this.destroyDeleteMaintenanceRunner();
    this.db.close();
  }

  async getSessionSummaries(limit = 50): Promise<SessionSummaryQueryRow[]> {
    return getSessionSummaries(this.db, limit);
  }

  async getSessionTimeline(sessionId: number, limit?: number): Promise<SessionTimelineRow[]> {
    return getSessionTimeline(this.db, sessionId, limit);
  }

  async getSessionWordsByLine(
    sessionId: number,
  ): Promise<Array<{ lineIndex: number; headword: string; occurrenceCount: number }>> {
    return getSessionWordsByLine(this.db, sessionId);
  }

  async getAllDistinctHeadwords(): Promise<string[]> {
    return getAllDistinctHeadwords(this.db);
  }

  async getAnimeDistinctHeadwords(animeId: number): Promise<string[]> {
    return getAnimeDistinctHeadwords(this.db, animeId);
  }

  async getMediaDistinctHeadwords(videoId: number): Promise<string[]> {
    return getMediaDistinctHeadwords(this.db, videoId);
  }

  async getQueryHints(): Promise<{
    totalSessions: number;
    activeSessions: number;
    episodesToday: number;
    activeAnimeCount: number;
    totalEpisodesWatched: number;
    totalAnimeCompleted: number;
    totalActiveMin: number;
    totalCards: number;
    activeDays: number;
    totalTokensSeen: number;
    totalLookupCount: number;
    totalLookupHits: number;
    totalYomitanLookupCount: number;
    newWordsToday: number;
    newWordsThisWeek: number;
  }> {
    return getQueryHints(this.db);
  }

  async getDailyRollups(limit = 60): Promise<ImmersionSessionRollupRow[]> {
    return getDailyRollups(this.db, limit);
  }

  async getMonthlyRollups(limit = 24): Promise<ImmersionSessionRollupRow[]> {
    return getMonthlyRollups(this.db, limit);
  }

  async getTrendsDashboard(
    range: '7d' | '30d' | '90d' | '365d' | 'all' = '30d',
    groupBy: 'day' | 'month' = 'day',
    fillEmptyBuckets = true,
  ) {
    return getTrendsDashboard(this.db, range, groupBy, fillEmptyBuckets);
  }

  async getVocabularyStats(limit = 100, excludePos?: string[]): Promise<VocabularyStatsRow[]> {
    return getVocabularyStats(this.db, limit, excludePos);
  }

  async getStatsExcludedWords(): Promise<StatsExcludedWordRow[]> {
    return getStatsExcludedWords(this.db);
  }

  async replaceStatsExcludedWords(words: StatsExcludedWordRow[]): Promise<void> {
    replaceStatsExcludedWords(this.db, words);
  }

  async cleanupVocabularyStats(): Promise<VocabularyCleanupSummary> {
    return cleanupVocabularyStats(this.db, {
      resolveLegacyPos: this.resolveLegacyVocabularyPos,
    });
  }

  /**
   * Collapse animation bursts that earlier versions recorded frame by frame. The whole
   * queue is drained first so a burst still waiting to be written is scanned as stored
   * rows rather than surviving the cleanup and landing a moment after it.
   */
  async cleanupDuplicateSubtitleLines(
    options: DuplicateSubtitleLineCleanupOptions = {},
  ): Promise<DuplicateSubtitleLineCleanupSummary> {
    this.drainQueue();
    return cleanupDuplicateSubtitleLines(this.db, options);
  }

  async rebuildLifetimeSummaries(): Promise<LifetimeRebuildSummary> {
    this.requireWriteQueueDrained('rebuilding lifetime summaries');
    // Non-destructive: recomputes from the media ledger (or bootstraps empty
    // lifetime tables), so history older than session retention is never reset.
    const repaired = repairLifetimeSummariesFromMedia(this.db);
    // Sessions currently tracked in the applied-sessions ledger, not sessions
    // processed by this call — the repair recomputes summaries instead of
    // re-applying sessions. Retention prunes these rows (FK cascade), so on an
    // old database this reads lower than the history the totals still include.
    const appliedRow = this.db
      .prepare('SELECT COUNT(*) AS count FROM imm_lifetime_applied_sessions')
      .get() as { count: number };
    return { appliedSessions: Number(appliedRow.count), rebuiltAtMs: repaired.repairedAtMs };
  }

  async getKanjiStats(limit = 100): Promise<KanjiStatsRow[]> {
    return getKanjiStats(this.db, limit);
  }

  async getWordOccurrences(
    headword: string,
    word: string,
    reading: string,
    limit = 100,
    offset = 0,
  ): Promise<WordOccurrenceRow[]> {
    return getWordOccurrences(this.db, headword, word, reading, limit, offset);
  }

  async getKanjiOccurrences(kanji: string, limit = 100, offset = 0): Promise<KanjiOccurrenceRow[]> {
    return getKanjiOccurrences(this.db, kanji, limit, offset);
  }

  async searchSubtitleSentences(
    query: string,
    limit = 50,
    options?: SentenceSearchOptions,
  ): Promise<SentenceSearchResultRow[]> {
    return searchSubtitleSentences(this.db, query, limit, options);
  }

  async getSessionEvents(
    sessionId: number,
    limit = 500,
    eventTypes?: number[],
  ): Promise<SessionEventRow[]> {
    return getSessionEvents(this.db, sessionId, limit, eventTypes);
  }

  async getMediaLibrary(): Promise<MediaLibraryRow[]> {
    const rows = getMediaLibrary(this.db);
    this.backfillYoutubeMetadataForLibrary();
    return rows;
  }

  async getMediaDetail(videoId: number): Promise<MediaDetailRow | null> {
    const detail = getMediaDetail(this.db, videoId);
    this.backfillYoutubeMetadataForVideo(videoId);
    return detail;
  }

  async getMediaSessions(videoId: number, limit = 100): Promise<SessionSummaryQueryRow[]> {
    return getMediaSessions(this.db, videoId, limit);
  }

  async getMediaDailyRollups(videoId: number, limit = 90): Promise<ImmersionSessionRollupRow[]> {
    return getMediaDailyRollups(this.db, videoId, limit);
  }

  async getCoverArt(videoId: number): Promise<MediaArtRow | null> {
    return getCoverArt(this.db, videoId);
  }

  async getAnimeLibrary(): Promise<AnimeLibraryRow[]> {
    this.relinkYoutubeAnimeLibrary();
    return getAnimeLibrary(this.db);
  }

  async getAnimeMergeRecommendations(): Promise<AnimeMergeRecommendation[]> {
    return getAnimeMergeRecommendations(this.db);
  }

  async dismissAnimeMergeRecommendation(recommendationId: number): Promise<boolean> {
    return dismissAnimeMergeRecommendation(this.db, recommendationId);
  }

  async getAnimeDetail(animeId: number): Promise<AnimeDetailRow | null> {
    this.relinkYoutubeAnimeLibrary();
    return getAnimeDetail(this.db, animeId);
  }

  async getAnimeEpisodes(animeId: number): Promise<AnimeEpisodeRow[]> {
    return getAnimeEpisodes(this.db, animeId);
  }

  async getAnimeAnilistEntries(animeId: number): Promise<AnimeAnilistEntryRow[]> {
    return getAnimeAnilistEntries(this.db, animeId);
  }

  async getAnimeCoverArt(animeId: number): Promise<MediaArtRow | null> {
    return getAnimeCoverArt(this.db, animeId);
  }

  async getAnimeWords(animeId: number, limit = 50): Promise<AnimeWordRow[]> {
    return getAnimeWords(this.db, animeId, limit);
  }

  async getEpisodeWords(videoId: number, limit = 50): Promise<AnimeWordRow[]> {
    return getEpisodeWords(this.db, videoId, limit);
  }

  async getEpisodeSessions(videoId: number): Promise<SessionSummaryQueryRow[]> {
    return getEpisodeSessions(this.db, videoId);
  }

  async setVideoWatched(videoId: number, watched: boolean): Promise<void> {
    markVideoWatched(this.db, videoId, watched);
  }

  /**
   * Set the watch mark on streamed episodes by hand.
   *
   * Marking watched creates the video row when the episode was never played, so
   * a series watched elsewhere can be caught up on; the row carries the same
   * series/season/episode metadata playback would have recorded. Both library
   * views join the lifetime tables, so a row created this way stays out of the
   * stats lists until it is actually watched.
   *
   * Clearing a mark never creates anything: with no row there is nothing to
   * clear.
   */
  async setStreamWatchState(
    episodes: StreamPlaybackMetadataInput[],
    watched: boolean,
  ): Promise<number> {
    let changed = 0;
    // Every row this creates would otherwise rebuild the lifetime summaries on
    // its own, and a season's worth of episodes arrives in one call.
    let needsLifetimeRebuild = false;
    // A half-applied batch would leave rows created without their marks, so the
    // whole span of episodes and the summary rebuild land together or not at all.
    let clearedActiveVideo = false;

    this.db.exec('BEGIN IMMEDIATE');
    try {
      for (const episode of episodes) {
        const statsPath = normalizeMediaPath(episode.statsPath);
        if (!statsPath) continue;
        if (watched) {
          needsLifetimeRebuild =
            this.recordStreamPlaybackMetadata(episode, { deferLifetimeRebuild: true }) ||
            needsLifetimeRebuild;
        }

        const videoId = getVideoIdByVideoKey(this.db, buildVideoKey(statsPath, SOURCE_TYPE_REMOTE));
        if (videoId === null) continue;
        markVideoWatched(this.db, videoId, watched);
        changed += 1;

        if (!watched && this.sessionState?.videoId === videoId) clearedActiveVideo = true;
      }

      if (needsLifetimeRebuild) rebuildLifetimeSummariesInTransaction(this.db);
      this.db.exec('COMMIT');
    } catch (error) {
      this.db.exec('ROLLBACK');
      throw error;
    }

    // Clearing the mark on what is playing right now would otherwise be undone
    // the moment the session passes the completion threshold again. Set after the
    // commit, so a rolled back clear does not suppress the automatic mark.
    if (clearedActiveVideo && this.sessionState) this.sessionState.markedWatched = true;
    return changed;
  }

  /**
   * Watch state for streamed episodes, keyed by the stats path the anime
   * browser derives for each one. Paths never played are absent from the map,
   * so a caller can treat "missing" as unwatched without a probe per episode.
   */
  async getStreamWatchState(statsPaths: string[]): Promise<Map<string, StreamWatchState>> {
    const byKey = new Map<string, string>();
    for (const path of statsPaths) {
      const normalized = normalizeMediaPath(path);
      if (!normalized) continue;
      byKey.set(buildVideoKey(normalized, SOURCE_TYPE_REMOTE), normalized);
    }

    const state = new Map<string, StreamWatchState>();
    for (const row of getWatchStateByVideoKeys(this.db, [...byKey.keys()])) {
      const statsPath = byKey.get(row.videoKey);
      if (!statsPath) continue;
      state.set(statsPath, {
        watched: row.watched,
        lastWatchedMs: row.lastWatchedMs,
        sessionCount: row.sessionCount,
      });
    }
    return state;
  }

  async markActiveVideoWatched(): Promise<boolean> {
    if (!this.sessionState) return false;
    markVideoWatched(this.db, this.sessionState.videoId, true);
    this.sessionState.markedWatched = true;
    return true;
  }

  async deleteSession(sessionId: number): Promise<void> {
    if (this.sessionState?.sessionId === sessionId) {
      this.logger.warn(`Ignoring delete request for active immersion session ${sessionId}`);
      return;
    }
    await this.enqueueDeleteMaintenanceTask(() => ({ kind: 'session', sessionId }));
  }

  async deleteSessions(sessionIds: number[]): Promise<void> {
    await this.enqueueDeleteMaintenanceTask(() => {
      const activeSessionId = this.sessionState?.sessionId;
      const deletableSessionIds =
        activeSessionId === undefined
          ? sessionIds
          : sessionIds.filter((sessionId) => sessionId !== activeSessionId);
      if (deletableSessionIds.length !== sessionIds.length) {
        this.logger.warn(
          `Ignoring bulk delete request for active immersion session ${activeSessionId}`,
        );
      }
      if (deletableSessionIds.length === 0) return null;
      return { kind: 'sessions', sessionIds: deletableSessionIds };
    });
  }

  async deleteVideo(videoId: number): Promise<void> {
    await this.enqueueDeleteMaintenanceTask(() => {
      if (this.sessionState?.videoId === videoId) {
        this.logger.warn(`Ignoring delete request for active immersion video ${videoId}`);
        return null;
      }
      return { kind: 'video', videoId };
    });
  }

  async deleteAnime(animeId: number): Promise<void> {
    await this.enqueueDeleteMaintenanceTask(async () => {
      // Resolve this at dispatch time because another queued delete can leave
      // enough time for playback to switch to an episode of this anime.
      const pendingVideoId = this.sessionState?.videoId;
      if (pendingVideoId !== undefined) {
        await this.pendingAnimeMetadataUpdates.get(pendingVideoId);
      }

      const activeVideoId = this.sessionState?.videoId;
      if (activeVideoId !== undefined) {
        const activeAnime = this.db
          .prepare('SELECT anime_id FROM imm_videos WHERE video_id = ?')
          .get(activeVideoId) as { anime_id: number | null } | null;
        if (activeAnime?.anime_id === animeId) {
          this.logger.warn(`Ignoring delete request for active immersion anime ${animeId}`);
          return null;
        }
      }
      return { kind: 'anime', animeId };
    });
  }

  private enqueueDeleteMaintenanceTask(
    resolveTask: Parameters<DeleteMaintenanceScheduler['enqueue']>[0],
  ): Promise<void> {
    if (this.isDestroyed) {
      return Promise.reject(new Error('Immersion tracker is shutting down'));
    }
    return this.deleteMaintenanceScheduler.enqueue(resolveTask);
  }

  /**
   * Fold duplicate library entries into one. Sources that hold the currently
   * playing episode are fine: the videos move, nothing is deleted out from
   * under the active session.
   */
  async mergeAnime(targetAnimeId: number, sourceAnimeIds: number[]): Promise<AnimeMergeSummary> {
    const pendingVideoId = this.sessionState?.videoId;
    if (pendingVideoId !== undefined) {
      await this.pendingAnimeMetadataUpdates.get(pendingVideoId);
    }
    // This rebuilds the lifetime summaries, which recompute from the database:
    // queued writes have to land first or the active session is dropped from
    // the merged totals.
    this.requireWriteQueueDrained('merging library entries');
    return mergeAnimeRecords(this.db, targetAnimeId, sourceAnimeIds);
  }

  async moveVideoToAnime(videoId: number, targetAnimeId: number): Promise<VideoMoveSummary> {
    await this.pendingAnimeMetadataUpdates.get(videoId);
    this.requireWriteQueueDrained('moving an episode');
    return moveVideoToAnimeQuery(this.db, videoId, targetAnimeId);
  }

  /**
   * Persist every queued write before a caller recomputes summaries from the
   * database.
   *
   * A single `flushNow()` is not enough: forced telemetry is appended to the
   * back of the queue while `flushNow()` writes at most `batchSize` entries off
   * the front, so a busy session leaves the newest sample unwritten. Stops as
   * soon as a pass makes no progress — a rolled-back batch is pushed back onto
   * the queue, and looping on that would spin forever.
   *
   * Returns false when the queue could not be emptied. Summary-rebuilding
   * callers fail closed in that case.
   */
  private drainWriteQueue(context: string): boolean {
    this.flushTelemetry(true);
    while (this.queue.length > 0) {
      const pending = this.queue.length;
      this.flushNow();
      if (this.queue.length >= pending) {
        this.logger.warn(
          `Immersion tracker queue did not drain before ${context}; summaries may lag by ${this.queue.length} writes`,
        );
        return false;
      }
    }
    return true;
  }

  private requireWriteQueueDrained(context: string): void {
    if (!this.drainWriteQueue(context)) {
      throw new Error(`Immersion tracker queue did not drain before ${context}`);
    }
  }

  async reassignAnimeAnilist(
    animeId: number,
    info: {
      anilistId: number;
      titleRomaji?: string | null;
      titleEnglish?: string | null;
      titleNative?: string | null;
      episodesTotal?: number | null;
      description?: string | null;
      coverUrl?: string | null;
    },
  ): Promise<void> {
    this.requireWriteQueueDrained('reassigning an AniList entry');
    // The user is acting on this entry, so it is the one that survives when
    // another row already claims the same AniList id.
    const repair = resolveAnimeAnilistConflict(this.db, animeId, info.anilistId, {
      survivor: 'target',
      matchConfidence: 'manual',
    });
    if (repair.anilistAssignmentBlocked) return;
    this.db
      .prepare(
        `
      UPDATE imm_anime
      SET anilist_id = ?,
          title_romaji = COALESCE(?, title_romaji),
          title_english = COALESCE(?, title_english),
          title_native = COALESCE(?, title_native),
          episodes_total = COALESCE(?, episodes_total),
          description = CASE WHEN ? = 1 THEN ? ELSE description END,
          LAST_UPDATE_DATE = ?
      WHERE anime_id = ?
    `,
      )
      .run(
        info.anilistId,
        info.titleRomaji ?? null,
        info.titleEnglish ?? null,
        info.titleNative ?? null,
        info.episodesTotal ?? null,
        info.description !== undefined ? 1 : 0,
        info.description ?? null,
        nowMs(),
        animeId,
      );
    // Empty lifetime tables still need the retained-session bootstrap. Once a
    // media ledger exists, only the redistributed and explicitly edited anime
    // can have changed.
    if (shouldBackfillLifetimeSummaries(this.db)) {
      repairLifetimeSummariesFromMedia(this.db);
    } else {
      const affectedAnimeIds = new Set(repair.affectedAnimeIds);
      affectedAnimeIds.add(animeId);
      recomputeLifetimeAnimeFromMedia(this.db, [...affectedAnimeIds]);
      recomputeLifetimeGlobalFromSummaries(this.db);
    }

    // Update cover art for all videos in this anime
    if (info.coverUrl) {
      const videos = this.db
        .prepare('SELECT video_id FROM imm_videos WHERE anime_id = ?')
        .all(animeId) as Array<{ video_id: number }>;
      let coverBlob: Buffer | null = null;
      try {
        const res = await fetch(info.coverUrl);
        if (res.ok) {
          coverBlob = Buffer.from(await res.arrayBuffer());
        }
      } catch {
        /* ignore */
      }
      for (const v of videos) {
        upsertCoverArt(this.db, v.video_id, {
          anilistId: info.anilistId,
          coverUrl: info.coverUrl,
          coverBlob,
          titleRomaji: info.titleRomaji ?? null,
          titleEnglish: info.titleEnglish ?? null,
          episodesTotal: info.episodesTotal ?? null,
        });
      }
    }
  }

  async getEpisodeCardEvents(videoId: number): Promise<EpisodeCardEventRow[]> {
    return getEpisodeCardEvents(this.db, videoId);
  }

  async getAnimeDailyRollups(animeId: number, limit = 90): Promise<ImmersionSessionRollupRow[]> {
    return getAnimeDailyRollups(this.db, animeId, limit);
  }

  async getStreakCalendar(days = 90): Promise<StreakCalendarRow[]> {
    return getStreakCalendar(this.db, days);
  }

  async getEpisodesPerDay(limit = 90): Promise<EpisodesPerDayRow[]> {
    return getEpisodesPerDay(this.db, limit);
  }

  async getNewAnimePerDay(limit = 90): Promise<NewAnimePerDayRow[]> {
    return getNewAnimePerDay(this.db, limit);
  }

  async getWatchTimePerAnime(limit = 90): Promise<WatchTimePerAnimeRow[]> {
    return getWatchTimePerAnime(this.db, limit);
  }

  async getWordDetail(wordId: number): Promise<WordDetailRow | null> {
    return getWordDetail(this.db, wordId);
  }

  async getWordAnimeAppearances(wordId: number): Promise<WordAnimeAppearanceRow[]> {
    return getWordAnimeAppearances(this.db, wordId);
  }

  async getSimilarWords(wordId: number, limit = 10): Promise<SimilarWordRow[]> {
    return getSimilarWords(this.db, wordId, limit);
  }

  async getKanjiDetail(kanjiId: number): Promise<KanjiDetailRow | null> {
    return getKanjiDetail(this.db, kanjiId);
  }

  async getKanjiAnimeAppearances(kanjiId: number): Promise<KanjiAnimeAppearanceRow[]> {
    return getKanjiAnimeAppearances(this.db, kanjiId);
  }

  async getKanjiWords(kanjiId: number, limit = 20): Promise<KanjiWordRow[]> {
    return getKanjiWords(this.db, kanjiId, limit);
  }

  setCoverArtFetcher(fetcher: CoverArtFetcher | null): void {
    this.coverArtFetcher = fetcher;
  }

  async ensureAnimeCoverArt(animeId: number): Promise<boolean> {
    const existing = await this.getAnimeCoverArt(animeId);
    if (existing?.coverBlob) {
      return true;
    }
    const row = this.db
      .prepare(
        'SELECT video_id AS videoId FROM imm_videos WHERE anime_id = ? ORDER BY video_id DESC LIMIT 1',
      )
      .get(animeId) as { videoId: number } | undefined;
    if (!row?.videoId) {
      return false;
    }
    return this.ensureCoverArt(row.videoId);
  }

  async ensureCoverArt(videoId: number): Promise<boolean> {
    const existing = await this.getCoverArt(videoId);
    if (existing?.coverBlob) {
      return true;
    }

    const row = this.db
      .prepare('SELECT source_url AS sourceUrl FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { sourceUrl: string | null } | null;
    const sourceUrl = row?.sourceUrl?.trim() ?? '';
    const youtubeVideoId = sourceUrl ? extractYouTubeVideoId(sourceUrl) : null;
    if (youtubeVideoId) {
      const youtubePromise = this.ensureYouTubeCoverArt(videoId, sourceUrl, youtubeVideoId);
      return await youtubePromise;
    }

    if (!this.coverArtFetcher) {
      return false;
    }
    const inFlight = this.pendingCoverFetches.get(videoId);
    if (inFlight) {
      return await inFlight;
    }

    const fetchPromise = (async () => {
      const titleRow = this.db
        .prepare('SELECT canonical_title AS canonicalTitle FROM imm_videos WHERE video_id = ?')
        .get(videoId) as { canonicalTitle: string | null } | undefined;
      const canonicalTitle = titleRow?.canonicalTitle?.trim();
      if (!canonicalTitle) {
        return false;
      }
      const fetched = await this.coverArtFetcher!.fetchIfMissing(this.db, videoId, canonicalTitle);
      if (!fetched) {
        return false;
      }
      const cover = await this.getCoverArt(videoId);
      return cover?.coverBlob != null;
    })();

    this.pendingCoverFetches.set(videoId, fetchPromise);
    try {
      return await fetchPromise;
    } finally {
      this.pendingCoverFetches.delete(videoId);
    }
  }

  private ensureYouTubeCoverArt(
    videoId: number,
    sourceUrl: string,
    youtubeVideoId: string,
  ): Promise<boolean> {
    const existing = this.pendingCoverFetches.get(videoId);
    if (existing) {
      return existing;
    }
    const promise = this.captureYouTubeCoverArt(videoId, sourceUrl, youtubeVideoId);
    this.pendingCoverFetches.set(videoId, promise);
    promise.finally(() => {
      this.pendingCoverFetches.delete(videoId);
    });
    return promise;
  }

  private async captureYouTubeCoverArt(
    videoId: number,
    sourceUrl: string,
    youtubeVideoId: string,
  ): Promise<boolean> {
    if (this.isDestroyed) return false;
    const existing = await this.getCoverArt(videoId);
    if (existing?.coverBlob) {
      return true;
    }
    if (
      existing?.coverUrl === null &&
      existing?.anilistId === null &&
      existing?.coverBlob === null &&
      nowMs() - existing.fetchedAtMs < YOUTUBE_COVER_RETRY_MS
    ) {
      return false;
    }

    let coverBlob: Buffer | null = null;
    let coverUrl: string | null = null;

    const embedThumbnailUrl = await fetchYouTubeOEmbedThumbnail(sourceUrl);
    if (embedThumbnailUrl) {
      const embedBlob = await downloadImage(embedThumbnailUrl);
      if (embedBlob) {
        coverBlob = embedBlob;
        coverUrl = embedThumbnailUrl;
      }
    }

    if (!coverBlob) {
      for (const candidate of buildYouTubeThumbnailUrls(youtubeVideoId)) {
        const candidateBlob = await downloadImage(candidate);
        if (!candidateBlob) {
          continue;
        }
        coverBlob = candidateBlob;
        coverUrl = candidate;
        break;
      }
    }

    if (!coverBlob) {
      const durationMs = getVideoDurationMs(this.db, videoId);
      const maxSeconds =
        durationMs > 0 ? Math.min(durationMs / 1000, YOUTUBE_SCREENSHOT_MAX_SECONDS) : null;
      const seekSecond = Math.random() * (maxSeconds ?? YOUTUBE_SCREENSHOT_MAX_SECONDS);
      try {
        coverBlob = await this.mediaGenerator.generateScreenshot(sourceUrl, seekSecond, {
          format: 'jpg',
          quality: 90,
          maxWidth: 640,
        });
      } catch (error) {
        this.logger.warn(
          'cover-art: failed to generate YouTube screenshot for videoId=%d: %s',
          videoId,
          (error as Error).message,
        );
      }
    }

    if (coverBlob) {
      upsertCoverArt(this.db, videoId, {
        anilistId: existing?.anilistId ?? null,
        coverUrl,
        coverBlob,
        titleRomaji: existing?.titleRomaji ?? null,
        titleEnglish: existing?.titleEnglish ?? null,
        episodesTotal: existing?.episodesTotal ?? null,
      });
      return true;
    }

    const shouldCacheNoMatch =
      !existing || (existing.coverUrl === null && existing.anilistId === null);
    if (shouldCacheNoMatch) {
      upsertCoverArt(this.db, videoId, {
        anilistId: null,
        coverUrl: null,
        coverBlob: null,
        titleRomaji: existing?.titleRomaji ?? null,
        titleEnglish: existing?.titleEnglish ?? null,
        episodesTotal: existing?.episodesTotal ?? null,
      });
    }

    return false;
  }

  private captureYoutubeMetadataAsync(videoId: number, sourceUrl: string): void {
    if (this.pendingYoutubeMetadataFetches.has(videoId)) {
      return;
    }

    const pending = (async () => {
      try {
        const metadata = await probeYoutubeVideoMetadata(sourceUrl);
        if (!metadata) {
          return;
        }
        upsertYoutubeVideoMetadata(this.db, videoId, metadata);
        linkYoutubeVideoToAnimeRecord(this.db, videoId, metadata);
        if (metadata.videoTitle?.trim()) {
          updateVideoTitleRecord(this.db, videoId, metadata.videoTitle.trim());
        }
      } catch (error) {
        this.logger.debug(
          'youtube metadata capture skipped for videoId=%d: %s',
          videoId,
          (error as Error).message,
        );
      }
    })();

    this.pendingYoutubeMetadataFetches.set(videoId, pending);
    pending.finally(() => {
      this.pendingYoutubeMetadataFetches.delete(videoId);
    });
  }

  private backfillYoutubeMetadataForLibrary(): void {
    const candidate = this.db
      .prepare(
        `
          SELECT
            v.video_id AS videoId,
            v.source_url AS sourceUrl
          FROM imm_videos v
          JOIN imm_lifetime_media lm ON lm.video_id = v.video_id
          LEFT JOIN imm_youtube_videos yv ON yv.video_id = v.video_id
          WHERE
            v.source_type = ?
            AND v.source_url IS NOT NULL
            AND (
              LOWER(v.source_url) LIKE 'https://www.youtube.com/%'
              OR LOWER(v.source_url) LIKE 'https://youtube.com/%'
              OR LOWER(v.source_url) LIKE 'https://m.youtube.com/%'
              OR LOWER(v.source_url) LIKE 'https://youtu.be/%'
            )
            AND (
              yv.video_id IS NULL
              OR yv.video_title IS NULL
              OR yv.channel_name IS NULL
              OR yv.channel_thumbnail_url IS NULL
            )
            AND (
              yv.fetched_at_ms IS NULL
              OR yv.fetched_at_ms <= ?
            )
          ORDER BY lm.last_watched_ms DESC, v.video_id DESC
          LIMIT 1
        `,
      )
      .get(SOURCE_TYPE_REMOTE, nowMs() - YOUTUBE_METADATA_REFRESH_MS) as {
      videoId: number;
      sourceUrl: string | null;
    } | null;
    if (!candidate?.sourceUrl) {
      return;
    }
    this.captureYoutubeMetadataAsync(candidate.videoId, candidate.sourceUrl);
  }

  private backfillYoutubeMetadataForVideo(videoId: number): void {
    const candidate = this.db
      .prepare(
        `
          SELECT
            v.source_url AS sourceUrl
          FROM imm_videos v
          LEFT JOIN imm_youtube_videos yv ON yv.video_id = v.video_id
          WHERE
            v.video_id = ?
            AND v.source_type = ?
            AND v.source_url IS NOT NULL
            AND (
              LOWER(v.source_url) LIKE 'https://www.youtube.com/%'
              OR LOWER(v.source_url) LIKE 'https://youtube.com/%'
              OR LOWER(v.source_url) LIKE 'https://m.youtube.com/%'
              OR LOWER(v.source_url) LIKE 'https://youtu.be/%'
            )
            AND (
              yv.video_id IS NULL
              OR yv.video_title IS NULL
              OR yv.channel_name IS NULL
              OR yv.channel_thumbnail_url IS NULL
            )
            AND (
              yv.fetched_at_ms IS NULL
              OR yv.fetched_at_ms <= ?
            )
        `,
      )
      .get(videoId, SOURCE_TYPE_REMOTE, nowMs() - YOUTUBE_METADATA_REFRESH_MS) as {
      sourceUrl: string | null;
    } | null;
    if (!candidate?.sourceUrl) {
      return;
    }
    this.captureYoutubeMetadataAsync(videoId, candidate.sourceUrl);
  }

  private relinkYoutubeAnimeLibrary(): void {
    const candidates = this.db
      .prepare(
        `
          SELECT
            v.video_id AS videoId,
            yv.youtube_video_id AS youtubeVideoId,
            yv.video_url AS videoUrl,
            yv.video_title AS videoTitle,
            yv.video_thumbnail_url AS videoThumbnailUrl,
            yv.channel_id AS channelId,
            yv.channel_name AS channelName,
            yv.channel_url AS channelUrl,
            yv.channel_thumbnail_url AS channelThumbnailUrl,
            yv.uploader_id AS uploaderId,
            yv.uploader_url AS uploaderUrl,
            yv.description AS description,
            yv.metadata_json AS metadataJson
          FROM imm_videos v
          JOIN imm_youtube_videos yv ON yv.video_id = v.video_id
          LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
          LEFT JOIN imm_lifetime_media lm ON lm.video_id = v.video_id
          WHERE
            v.source_type = ?
            AND v.anime_assignment_locked = 0
            AND v.source_url IS NOT NULL
            AND (
              LOWER(v.source_url) LIKE 'https://www.youtube.com/%'
              OR LOWER(v.source_url) LIKE 'https://youtube.com/%'
              OR LOWER(v.source_url) LIKE 'https://m.youtube.com/%'
              OR LOWER(v.source_url) LIKE 'https://youtu.be/%'
            )
            AND yv.channel_name IS NOT NULL
            AND (
              v.anime_id IS NULL
              OR a.metadata_json IS NULL
              OR a.metadata_json NOT LIKE '%"source":"youtube-channel"%'
              OR a.canonical_title IS NULL
              OR TRIM(a.canonical_title) != TRIM(yv.channel_name)
            )
          ORDER BY lm.last_watched_ms DESC, v.video_id DESC
        `,
      )
      .all(SOURCE_TYPE_REMOTE) as Array<{
      videoId: number;
      youtubeVideoId: string | null;
      videoUrl: string | null;
      videoTitle: string | null;
      videoThumbnailUrl: string | null;
      channelId: string | null;
      channelName: string | null;
      channelUrl: string | null;
      channelThumbnailUrl: string | null;
      uploaderId: string | null;
      uploaderUrl: string | null;
      description: string | null;
      metadataJson: string | null;
    }>;

    if (candidates.length === 0) {
      return;
    }

    for (const candidate of candidates) {
      if (!candidate.youtubeVideoId || !candidate.videoUrl) {
        continue;
      }
      linkYoutubeVideoToAnimeRecord(this.db, candidate.videoId, {
        youtubeVideoId: candidate.youtubeVideoId,
        videoUrl: candidate.videoUrl,
        videoTitle: candidate.videoTitle,
        videoThumbnailUrl: candidate.videoThumbnailUrl,
        channelId: candidate.channelId,
        channelName: candidate.channelName,
        channelUrl: candidate.channelUrl,
        channelThumbnailUrl: candidate.channelThumbnailUrl,
        uploaderId: candidate.uploaderId,
        uploaderUrl: candidate.uploaderUrl,
        description: candidate.description,
        metadataJson: candidate.metadataJson,
      });
    }
    repairLifetimeSummariesFromMedia(this.db);
  }

  recordJellyfinPlaybackMetadata(metadata: JellyfinPlaybackMetadataInput): void {
    const rawPath = normalizeMediaPath(metadata.mediaPath);
    if (!rawPath) {
      return;
    }
    const statsPath = buildJellyfinStatsMediaPath(rawPath, metadata.itemId);
    const displayTitle =
      normalizeText(metadata.displayTitle) ||
      normalizeText(metadata.itemTitle) ||
      deriveCanonicalTitle(statsPath);
    const itemTitle = normalizeText(metadata.itemTitle) || displayTitle;
    const seriesTitle = normalizeText(metadata.seriesTitle);
    const libraryTitle = seriesTitle || itemTitle;
    const seasonNumber = normalizeMetadataInt(metadata.seasonNumber);
    const episodeNumber = normalizeMetadataInt(metadata.episodeNumber);
    if (!libraryTitle) {
      return;
    }

    this.recordPrePlaybackMetadata({
      statsPath,
      aliases: buildMediaPathAliasCandidates(rawPath),
      displayTitle,
      libraryTitle,
      seasonNumber,
      episodeNumber,
      parserSource: 'jellyfin',
      metadataJson: JSON.stringify({
        source: 'jellyfin',
        itemId: normalizeText(metadata.itemId) || null,
        itemTitle,
        seriesTitle: seriesTitle || null,
        displayTitle,
        seasonNumber,
        episodeNumber,
      }),
    });
  }

  /** Returns whether the lifetime summaries still need rebuilding; see
   * `recordPrePlaybackMetadata` for why a batch defers that. */
  recordStreamPlaybackMetadata(
    metadata: StreamPlaybackMetadataInput,
    options: { deferLifetimeRebuild?: boolean } = {},
  ): boolean {
    const rawPath = normalizeMediaPath(metadata.mediaPath);
    const statsPath = normalizeMediaPath(metadata.statsPath) || rawPath;
    if (!statsPath) {
      return false;
    }
    const seriesTitle = normalizeText(metadata.seriesTitle);
    const displayTitle =
      normalizeText(metadata.displayTitle) || seriesTitle || deriveCanonicalTitle(statsPath);
    const libraryTitle = seriesTitle || displayTitle;
    if (!libraryTitle) {
      return false;
    }
    const seasonNumber = normalizeMetadataInt(metadata.seasonNumber);
    const episodeNumber = normalizeMetadataInt(metadata.episodeNumber);

    return this.recordPrePlaybackMetadata({
      deferLifetimeRebuild: options.deferLifetimeRebuild,
      statsPath,
      aliases: rawPath ? buildMediaPathAliasCandidates(rawPath) : [],
      displayTitle,
      libraryTitle,
      seasonNumber,
      episodeNumber,
      parserSource: 'anime-browser',
      metadataJson: JSON.stringify({
        source: 'anime-browser',
        seriesTitle: seriesTitle || null,
        displayTitle,
        seasonNumber,
        episodeNumber,
      }),
    });
  }

  /**
   * Creates the video row and its series link ahead of playback, so the session
   * mpv's path change starts already belongs to the right anime.
   *
   * Returns whether the lifetime summaries need rebuilding. A new row always
   * does, so a caller recording a batch passes `deferLifetimeRebuild` and
   * rebuilds once at the end rather than once per episode.
   */
  private recordPrePlaybackMetadata(params: {
    statsPath: string;
    aliases: string[];
    displayTitle: string;
    libraryTitle: string;
    seasonNumber: number | null;
    episodeNumber: number | null;
    parserSource: string;
    metadataJson: string;
    deferLifetimeRebuild?: boolean;
  }): boolean {
    for (const alias of params.aliases) {
      this.mediaPathAliases.set(alias, params.statsPath);
    }

    const videoId = getOrCreateVideoRecord(
      this.db,
      buildVideoKey(params.statsPath, SOURCE_TYPE_REMOTE),
      {
        canonicalTitle: params.displayTitle,
        sourcePath: null,
        sourceUrl: params.statsPath,
        sourceType: SOURCE_TYPE_REMOTE,
      },
    );
    const previousLink = this.db
      .prepare('SELECT anime_id AS animeId FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { animeId: number | null } | null;
    const animeId =
      getManualAnimeAssignment(this.db, videoId) ??
      getOrCreateAnimeRecord(this.db, {
        parsedTitle: params.libraryTitle,
        canonicalTitle: params.libraryTitle,
        seasonScope: params.seasonNumber,
        anilistId: null,
        titleRomaji: null,
        titleEnglish: null,
        titleNative: null,
        metadataJson: params.metadataJson,
      });
    linkVideoToAnimeRecord(this.db, videoId, {
      animeId,
      parsedBasename: null,
      parsedTitle: params.libraryTitle,
      parsedSeason: params.seasonNumber,
      parsedEpisode: params.episodeNumber,
      parserSource: params.parserSource,
      parserConfidence: 1,
      parseMetadataJson: params.metadataJson,
    });

    const hasLifetimeMedia = Boolean(
      this.db.prepare('SELECT 1 FROM imm_lifetime_media WHERE video_id = ?').get(videoId),
    );
    const needsRebuild = Boolean(
      hasLifetimeMedia || (previousLink && previousLink.animeId !== animeId),
    );
    if (needsRebuild && !params.deferLifetimeRebuild) {
      // Playback-time relink: only the old and new anime are affected, so
      // recompute just those from the media ledger instead of a full repair.
      const affectedAnimeIds = new Set<number>([animeId]);
      if (previousLink?.animeId) affectedAnimeIds.add(previousLink.animeId);
      let transactionStarted = false;
      try {
        this.db.exec('BEGIN IMMEDIATE');
        transactionStarted = true;
        recomputeLifetimeAnimeFromMedia(this.db, [...affectedAnimeIds]);
        recomputeLifetimeGlobalFromSummaries(this.db);
        this.db.exec('COMMIT');
      } catch (error) {
        if (transactionStarted) this.db.exec('ROLLBACK');
        throw error;
      }
    }
    return needsRebuild;
  }

  private hasPrePlaybackMetadata(videoId: number): boolean {
    const row = this.db
      .prepare('SELECT parser_source AS parserSource FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { parserSource: string | null } | null;
    return row?.parserSource !== null && PREPLAYBACK_PARSER_SOURCES.has(row?.parserSource ?? '');
  }

  handleMediaChange(mediaPath: string | null, mediaTitle: string | null): void {
    const rawPath = normalizeMediaPath(mediaPath);
    const normalizedPath =
      buildMediaPathAliasCandidates(rawPath)
        .map((alias) => this.mediaPathAliases.get(alias))
        .find((alias): alias is string => Boolean(alias)) ?? rawPath;
    const normalizedTitle = normalizeText(mediaTitle);
    this.logger.info(
      `handleMediaChange called with path=${normalizedPath || '<empty>'} title=${normalizedTitle || '<empty>'}`,
    );
    if (normalizedPath === this.currentMediaPathOrUrl) {
      if (normalizedTitle && normalizedTitle !== this.currentVideoKey) {
        this.currentVideoKey = normalizedTitle;
        this.updateVideoTitleForActiveSession(normalizedTitle);
        this.logger.debug('Media title updated for existing session');
      } else {
        this.logger.debug('Media change ignored; path unchanged');
      }
      return;
    }
    this.finalizeActiveSession();
    this.currentMediaPathOrUrl = normalizedPath;
    this.currentVideoKey = normalizedTitle;
    if (!normalizedPath) {
      this.logger.info('Media path cleared; immersion session tracking paused');
      return;
    }

    const sourceType = isRemoteSource(normalizedPath) ? SOURCE_TYPE_REMOTE : SOURCE_TYPE_LOCAL;
    const videoKey = buildVideoKey(normalizedPath, sourceType);
    const canonicalTitle = normalizedTitle || deriveCanonicalTitle(normalizedPath);
    const sourcePath = sourceType === SOURCE_TYPE_LOCAL ? normalizedPath : null;
    const sourceUrl = sourceType === SOURCE_TYPE_REMOTE ? normalizedPath : null;

    const sessionInfo = {
      videoId: getOrCreateVideoRecord(this.db, videoKey, {
        canonicalTitle,
        sourcePath,
        sourceUrl,
        sourceType,
      }),
      startedAtMs: nowMs(),
    };

    this.logger.info(
      `Starting immersion session for path=${normalizedPath} videoId=${sessionInfo.videoId}`,
    );
    this.startSession(sessionInfo.videoId, sessionInfo.startedAtMs);
    const youtubeVideoId =
      sourceType === SOURCE_TYPE_REMOTE ? extractYouTubeVideoId(normalizedPath) : null;
    if (youtubeVideoId) {
      void this.ensureYouTubeCoverArt(sessionInfo.videoId, normalizedPath, youtubeVideoId);
      this.captureYoutubeMetadataAsync(sessionInfo.videoId, normalizedPath);
    } else if (!this.hasPrePlaybackMetadata(sessionInfo.videoId)) {
      this.captureAnimeMetadataAsync(sessionInfo.videoId, normalizedPath, normalizedTitle || null);
    }
    if (!youtubeVideoId) {
      this.prefetchCoverArtAsync(sessionInfo.videoId);
    }
    this.captureVideoMetadataAsync(sessionInfo.videoId, sourceType, normalizedPath);
  }

  handleMediaTitleUpdate(mediaTitle: string | null): void {
    if (!this.sessionState) return;
    const normalizedTitle = normalizeText(mediaTitle);
    if (!normalizedTitle) return;
    this.currentVideoKey = normalizedTitle;
    this.updateVideoTitleForActiveSession(normalizedTitle);
  }

  recordSubtitleLine(
    text: string,
    startSec: number,
    endSec: number,
    tokens?: MergedToken[] | null,
    secondaryText?: string | null,
  ): void {
    if (!this.sessionState || !text.trim()) return;
    const cleaned = normalizeText(text);
    if (!cleaned) return;

    if (!Number.isFinite(startSec) || !Number.isFinite(endSec) || endSec <= startSec) {
      return;
    }

    const startMs = secToMs(startSec);
    const endMs = secToMs(endSec);
    const subtitleKey = `${startMs}:${cleaned}`;
    if (this.recordedSubtitleKeys.has(subtitleKey)) {
      return;
    }
    this.recordedSubtitleKeys.add(subtitleKey);

    const currentTimeMs = nowMs();
    const nowSec = currentTimeMs / 1000;

    const tokenCount = tokens?.length ?? 0;
    this.sessionState.currentLineIndex += 1;
    this.sessionState.linesSeen += 1;
    this.sessionState.tokensSeen += tokenCount;
    if (this.sessionState.lastMediaMs === null || endMs > this.sessionState.lastMediaMs) {
      this.sessionState.lastMediaMs = endMs;
    }
    this.sessionState.pendingTelemetry = true;

    const wordOccurrences = new Map<string, CountedWordOccurrence>();
    for (const token of tokens ?? []) {
      if (shouldExcludeTokenFromVocabularyPersistence(token)) {
        continue;
      }
      const headword = normalizeText(token.headword || token.surface);
      const word = normalizeText(token.surface || token.headword);
      const reading = normalizeText(token.reading);
      if (!headword || !word) {
        continue;
      }
      const wordKey = [headword, word, reading].join('\u0000');
      const storedPartOfSpeech = deriveStoredPartOfSpeech({
        partOfSpeech: token.partOfSpeech,
        pos1: token.pos1 ?? '',
      });
      const existing = wordOccurrences.get(wordKey);
      if (existing) {
        existing.occurrenceCount += 1;
        continue;
      }
      wordOccurrences.set(wordKey, {
        headword,
        word,
        reading,
        partOfSpeech: storedPartOfSpeech,
        pos1: token.pos1 ?? '',
        pos2: token.pos2 ?? '',
        pos3: token.pos3 ?? '',
        occurrenceCount: 1,
        frequencyRank: token.frequencyRank ?? null,
      });
    }

    const kanjiCounts = new Map<string, number>();
    for (const char of cleaned) {
      if (!isKanji(char)) {
        continue;
      }
      kanjiCounts.set(char, (kanjiCounts.get(char) ?? 0) + 1);
    }

    this.recordWrite({
      kind: 'subtitleLine',
      sessionId: this.sessionState.sessionId,
      videoId: this.sessionState.videoId,
      lineIndex: this.sessionState.currentLineIndex,
      segmentStartMs: startMs,
      segmentEndMs: endMs,
      text: cleaned,
      secondaryText: secondaryText ?? null,
      wordOccurrences: Array.from(wordOccurrences.values()),
      kanjiOccurrences: Array.from(kanjiCounts.entries()).map(([kanji, occurrenceCount]) => ({
        kanji,
        occurrenceCount,
      })),
      firstSeen: nowSec,
      lastSeen: nowSec,
    });

    this.recordWrite({
      kind: 'event',
      sessionId: this.sessionState.sessionId,
      sampleMs: currentTimeMs,
      lineIndex: this.sessionState.currentLineIndex,
      segmentStartMs: secToMs(startSec),
      segmentEndMs: secToMs(endSec),
      tokensDelta: tokenCount,
      cardsDelta: 0,
      eventType: EVENT_SUBTITLE_LINE,
      payloadJson: sanitizePayload(
        {
          event: 'subtitle-line',
          tokens: tokenCount,
        },
        this.maxPayloadBytes,
      ),
    });
  }

  recordMediaDuration(durationSec: number): void {
    if (!this.sessionState || !Number.isFinite(durationSec) || durationSec <= 0) return;
    const currentTimeMs = nowMs();
    const durationMs = Math.round(durationSec * 1000);
    const current = getVideoDurationMs(this.db, this.sessionState.videoId);
    if (current === 0 || Math.abs(current - durationMs) > 1000) {
      this.db
        .prepare('UPDATE imm_videos SET duration_ms = ?, LAST_UPDATE_DATE = ? WHERE video_id = ?')
        .run(durationMs, currentTimeMs, this.sessionState.videoId);
    }
  }

  recordPlaybackPosition(mediaTimeSec: number | null): void {
    if (!this.sessionState || mediaTimeSec === null || !Number.isFinite(mediaTimeSec)) {
      return;
    }
    const currentTimeMs = nowMs();
    const mediaMs = Math.round(mediaTimeSec * 1000);
    if (this.sessionState.lastWallClockMs <= 0) {
      this.sessionState.lastWallClockMs = currentTimeMs;
      this.sessionState.lastMediaMs = mediaMs;
      return;
    }

    const wallDeltaMs = currentTimeMs - this.sessionState.lastWallClockMs;
    if (wallDeltaMs > 0 && wallDeltaMs < 60_000) {
      this.sessionState.totalWatchedMs += wallDeltaMs;
      if (!this.sessionState.isPaused) {
        this.sessionState.activeWatchedMs += wallDeltaMs;
      }
    }

    if (this.sessionState.lastMediaMs !== null) {
      const mediaDeltaMs = mediaMs - this.sessionState.lastMediaMs;
      if (Math.abs(mediaDeltaMs) >= 1_000) {
        if (mediaDeltaMs > 0) {
          this.sessionState.seekForwardCount += 1;
          this.sessionState.pendingTelemetry = true;
          this.recordWrite({
            kind: 'event',
            sessionId: this.sessionState.sessionId,
            sampleMs: currentTimeMs,
            eventType: EVENT_SEEK_FORWARD,
            tokensDelta: 0,
            cardsDelta: 0,
            segmentStartMs: this.sessionState.lastMediaMs,
            segmentEndMs: mediaMs,
            payloadJson: sanitizePayload(
              {
                fromMs: this.sessionState.lastMediaMs,
                toMs: mediaMs,
              },
              this.maxPayloadBytes,
            ),
          });
        } else if (mediaDeltaMs < 0) {
          this.sessionState.seekBackwardCount += 1;
          this.sessionState.pendingTelemetry = true;
          this.recordWrite({
            kind: 'event',
            sessionId: this.sessionState.sessionId,
            sampleMs: currentTimeMs,
            eventType: EVENT_SEEK_BACKWARD,
            tokensDelta: 0,
            cardsDelta: 0,
            segmentStartMs: this.sessionState.lastMediaMs,
            segmentEndMs: mediaMs,
            payloadJson: sanitizePayload(
              {
                fromMs: this.sessionState.lastMediaMs,
                toMs: mediaMs,
              },
              this.maxPayloadBytes,
            ),
          });
        }
      }
    }

    this.sessionState.lastWallClockMs = currentTimeMs;
    this.sessionState.lastMediaMs = mediaMs;
    this.sessionState.pendingTelemetry = true;

    if (!this.sessionState.markedWatched) {
      const durationMs = getVideoDurationMs(this.db, this.sessionState.videoId);
      if (durationMs > 0 && mediaMs >= durationMs * DEFAULT_MIN_WATCH_RATIO) {
        markVideoWatched(this.db, this.sessionState.videoId, true);
        this.sessionState.markedWatched = true;
      }
    }
  }

  recordPauseState(isPaused: boolean): void {
    if (!this.sessionState) return;
    if (this.sessionState.isPaused === isPaused) return;

    const currentTimeMs = nowMs();
    this.sessionState.isPaused = isPaused;
    if (isPaused) {
      this.sessionState.lastPauseStartMs = currentTimeMs;
      this.sessionState.pauseCount += 1;
      this.recordWrite({
        kind: 'event',
        sessionId: this.sessionState.sessionId,
        sampleMs: currentTimeMs,
        eventType: EVENT_PAUSE_START,
        cardsDelta: 0,
        tokensDelta: 0,
        payloadJson: sanitizePayload({ paused: true }, this.maxPayloadBytes),
      });
    } else {
      if (this.sessionState.lastPauseStartMs) {
        const pauseMs = Math.max(0, currentTimeMs - this.sessionState.lastPauseStartMs);
        this.sessionState.pauseMs += pauseMs;
        this.sessionState.lastPauseStartMs = null;
      }
      this.recordWrite({
        kind: 'event',
        sessionId: this.sessionState.sessionId,
        sampleMs: currentTimeMs,
        eventType: EVENT_PAUSE_END,
        cardsDelta: 0,
        tokensDelta: 0,
        payloadJson: sanitizePayload({ paused: false }, this.maxPayloadBytes),
      });
    }

    this.sessionState.pendingTelemetry = true;
  }

  recordLookup(hit: boolean): void {
    if (!this.sessionState) return;
    this.sessionState.lookupCount += 1;
    if (hit) {
      this.sessionState.lookupHits += 1;
    }
    this.sessionState.pendingTelemetry = true;
    this.recordWrite({
      kind: 'event',
      sessionId: this.sessionState.sessionId,
      sampleMs: nowMs(),
      eventType: EVENT_LOOKUP,
      cardsDelta: 0,
      tokensDelta: 0,
      payloadJson: sanitizePayload(
        {
          hit,
        },
        this.maxPayloadBytes,
      ),
    });
  }

  recordYomitanLookup(): void {
    if (!this.sessionState) return;
    this.sessionState.yomitanLookupCount += 1;
    this.sessionState.pendingTelemetry = true;
    this.recordWrite({
      kind: 'event',
      sessionId: this.sessionState.sessionId,
      sampleMs: nowMs(),
      eventType: EVENT_YOMITAN_LOOKUP,
      cardsDelta: 0,
      tokensDelta: 0,
      payloadJson: null,
    });
  }

  recordCardsMined(count = 1, noteIds?: number[]): void {
    if (!this.sessionState) return;
    this.sessionState.cardsMined += count;
    this.sessionState.pendingTelemetry = true;
    this.recordWrite({
      kind: 'event',
      sessionId: this.sessionState.sessionId,
      sampleMs: nowMs(),
      eventType: EVENT_CARD_MINED,
      tokensDelta: 0,
      cardsDelta: count,
      payloadJson: sanitizePayload(
        { cardsMined: count, ...(noteIds?.length ? { noteIds } : {}) },
        this.maxPayloadBytes,
      ),
    });
  }

  recordMediaBufferEvent(): void {
    if (!this.sessionState) return;
    this.sessionState.mediaBufferEvents += 1;
    this.sessionState.pendingTelemetry = true;
    this.recordWrite({
      kind: 'event',
      sessionId: this.sessionState.sessionId,
      sampleMs: nowMs(),
      eventType: EVENT_MEDIA_BUFFER,
      cardsDelta: 0,
      tokensDelta: 0,
      payloadJson: sanitizePayload(
        {
          buffer: true,
        },
        this.maxPayloadBytes,
      ),
    });
  }

  private recordWrite(write: QueuedWrite): void {
    if (this.isDestroyed) return;
    const { dropped } = enqueueWrite(this.queue, write, this.queueCap);
    if (dropped > 0) {
      this.droppedWriteCount += dropped;
      this.logger.warn(`Immersion tracker queue overflow; dropped ${dropped} oldest writes`);
    }
    if (write.kind === 'event' || this.queue.length >= this.batchSize) {
      this.scheduleFlush(0);
    }
  }

  private flushTelemetry(force = false): void {
    if (!this.sessionState || (!force && !this.sessionState.pendingTelemetry)) {
      return;
    }
    this.recordWrite({
      kind: 'telemetry',
      sessionId: this.sessionState.sessionId,
      sampleMs: nowMs(),
      lastMediaMs: this.sessionState.lastMediaMs,
      totalWatchedMs: this.sessionState.totalWatchedMs,
      activeWatchedMs: this.sessionState.activeWatchedMs,
      linesSeen: this.sessionState.linesSeen,
      tokensSeen: this.sessionState.tokensSeen,
      cardsMined: this.sessionState.cardsMined,
      lookupCount: this.sessionState.lookupCount,
      lookupHits: this.sessionState.lookupHits,
      yomitanLookupCount: this.sessionState.yomitanLookupCount,
      pauseCount: this.sessionState.pauseCount,
      pauseMs: this.sessionState.pauseMs,
      seekForwardCount: this.sessionState.seekForwardCount,
      seekBackwardCount: this.sessionState.seekBackwardCount,
      mediaBufferEvents: this.sessionState.mediaBufferEvents,
    });
    this.sessionState.pendingTelemetry = false;
  }

  private scheduleFlush(delayMs = this.flushIntervalMs): void {
    if (this.flushScheduled || this.writeLock.locked) return;
    this.flushScheduled = true;
    this.flushTimer = setTimeout(() => {
      this.flushScheduled = false;
      this.flushNow();
    }, delayMs);
  }

  private flushNow(): void {
    if (this.writeLock.locked || this.isDestroyed) return;
    if (this.queue.length === 0) {
      this.flushScheduled = false;
      return;
    }

    this.flushTelemetry();
    if (this.queue.length === 0) {
      this.flushScheduled = false;
      return;
    }

    const batch = this.queue.splice(0, Math.min(this.batchSize, this.queue.length));
    this.writeLock.locked = true;
    try {
      this.db.exec('BEGIN IMMEDIATE');
      for (const write of batch) {
        this.flushSingle(write);
      }
      this.db.exec('COMMIT');
    } catch (error) {
      this.db.exec('ROLLBACK');
      this.queue.unshift(...batch);
      this.logger.warn('Immersion tracker flush failed, retrying later', error as Error);
    } finally {
      this.writeLock.locked = false;
      this.flushScheduled = false;
      if (this.queue.length > 0) {
        this.scheduleFlush(this.flushIntervalMs);
      }
    }
  }

  /**
   * Write out everything queued, not just the next batch.
   *
   * `flushNow` writes at most `batchSize` entries and does nothing at all while the write
   * lock is held, so a maintenance pass that runs straight after it can still be reading
   * a database that is missing rows. Each pass has to shrink the queue to continue: a
   * failed flush puts its batch back, and looping on that would never finish.
   */
  private drainQueue(): void {
    while (this.queue.length > 0) {
      const pendingBefore = this.queue.length;
      this.flushNow();
      if (this.queue.length >= pendingBefore) {
        return;
      }
    }
  }

  private flushSingle(write: QueuedWrite): void {
    executeQueuedWrite(write, this.preparedStatements);
  }

  private scheduleMaintenance(): void {
    this.maintenanceTimer = setInterval(() => {
      this.runMaintenance();
    }, this.maintenanceIntervalMs);
    this.runMaintenance();
  }

  private runMaintenance(): void {
    if (this.isDestroyed || this.writeLock.locked) return;
    try {
      this.flushTelemetry(true);
      this.flushNow();
      const maintenanceNowMs = nowMs();
      this.runRollupMaintenance(false);
      if (
        Number.isFinite(this.eventsRetentionMs) ||
        Number.isFinite(this.telemetryRetentionMs) ||
        Number.isFinite(this.sessionsRetentionMs)
      ) {
        pruneRawRetention(this.db, maintenanceNowMs, {
          eventsRetentionMs: this.eventsRetentionMs,
          telemetryRetentionMs: this.telemetryRetentionMs,
          sessionsRetentionMs: this.sessionsRetentionMs,
          eventsRetentionDays: this.eventsRetentionDays ?? undefined,
          telemetryRetentionDays: this.telemetryRetentionDays ?? undefined,
          sessionsRetentionDays: this.sessionsRetentionDays ?? undefined,
        });
      }
      if (
        Number.isFinite(this.dailyRollupRetentionMs) ||
        Number.isFinite(this.monthlyRollupRetentionMs)
      ) {
        pruneRollupRetention(this.db, maintenanceNowMs, {
          dailyRollupRetentionMs: this.dailyRollupRetentionMs,
          monthlyRollupRetentionMs: this.monthlyRollupRetentionMs,
        });
      }

      if (
        this.vacuumIntervalMs > 0 &&
        maintenanceNowMs - this.lastVacuumMs >= this.vacuumIntervalMs &&
        !this.writeLock.locked
      ) {
        this.db.exec('VACUUM');
        this.lastVacuumMs = maintenanceNowMs;
      }
      runOptimizeMaintenance(this.db);
    } catch (error) {
      this.logger.warn(
        'Immersion tracker maintenance failed, will retry later',
        (error as Error).message,
      );
    }
  }

  private runRollupMaintenance(forceRebuild = false): void {
    runRollupMaintenance(this.db, forceRebuild);
  }

  private startSession(videoId: number, startedAtMs?: number): void {
    const { sessionId, state } = startSessionRecord(this.db, videoId, startedAtMs);
    this.sessionState = state;
    this.recordedSubtitleKeys.clear();
    this.recordWrite({
      kind: 'telemetry',
      sessionId,
      sampleMs: state.startedAtMs,
      totalWatchedMs: 0,
      activeWatchedMs: 0,
      linesSeen: 0,
      tokensSeen: 0,
      cardsMined: 0,
      lookupCount: 0,
      lookupHits: 0,
      yomitanLookupCount: 0,
      pauseCount: 0,
      pauseMs: 0,
      seekForwardCount: 0,
      seekBackwardCount: 0,
      mediaBufferEvents: 0,
    });
    this.scheduleFlush(0);
  }

  private finalizeActiveSession(): void {
    if (!this.sessionState) return;
    const endedAt = nowMs();
    if (this.sessionState.lastPauseStartMs) {
      this.sessionState.pauseMs += Math.max(0, endedAt - this.sessionState.lastPauseStartMs);
      this.sessionState.lastPauseStartMs = null;
    }
    const finalWallNow = endedAt;
    if (this.sessionState.lastWallClockMs > 0) {
      const wallDelta = finalWallNow - this.sessionState.lastWallClockMs;
      if (wallDelta > 0 && wallDelta < 60_000) {
        this.sessionState.totalWatchedMs += wallDelta;
        if (!this.sessionState.isPaused) {
          this.sessionState.activeWatchedMs += wallDelta;
        }
      }
    }
    this.flushTelemetry(true);
    this.flushNow();
    this.sessionState.pendingTelemetry = false;

    finalizeSessionRecord(this.db, this.sessionState, endedAt);
    applySessionLifetimeSummary(this.db, this.sessionState, endedAt);
    this.sessionState = null;
  }

  private captureVideoMetadataAsync(videoId: number, sourceType: number, mediaPath: string): void {
    if (sourceType !== SOURCE_TYPE_LOCAL) return;
    void (async () => {
      try {
        const metadata = await getLocalVideoMetadata(mediaPath);
        updateVideoMetadataRecord(this.db, videoId, metadata);
      } catch (error) {
        this.logger.warn('Unable to capture local video metadata', (error as Error).message);
      }
    })();
  }

  private captureAnimeMetadataAsync(
    videoId: number,
    mediaPath: string | null,
    mediaTitle: string | null,
  ): void {
    const updatePromise = (async () => {
      try {
        const parsed = await guessAnimeVideoMetadata(mediaPath, mediaTitle);
        if (this.isDestroyed || !parsed?.parsedTitle.trim()) {
          return;
        }

        const animeId =
          getManualAnimeAssignment(this.db, videoId) ??
          (mediaPath && !isRemoteSource(mediaPath)
            ? findManualDirectoryAnimeAssignment(
                this.db,
                videoId,
                mediaPath,
                parsed.parsedTitle,
                parsed.parsedSeason,
              )
            : null) ??
          getOrCreateAnimeRecord(this.db, {
            parsedTitle: parsed.parsedTitle,
            canonicalTitle: parsed.parsedTitle,
            seasonScope: parsed.parsedSeason,
            anilistId: null,
            titleRomaji: null,
            titleEnglish: null,
            titleNative: null,
            metadataJson: parsed.parseMetadataJson,
          });
        linkVideoToAnimeRecord(this.db, videoId, {
          animeId,
          parsedBasename: parsed.parsedBasename,
          parsedTitle: parsed.parsedTitle,
          parsedSeason: parsed.parsedSeason,
          parsedEpisode: parsed.parsedEpisode,
          parserSource: parsed.parserSource,
          parserConfidence: parsed.parserConfidence,
          parseMetadataJson: parsed.parseMetadataJson,
        });
      } catch (error) {
        this.logger.warn('Unable to capture anime metadata', (error as Error).message);
      }
    })();

    this.pendingAnimeMetadataUpdates.set(videoId, updatePromise);
    void updatePromise.finally(() => {
      this.pendingAnimeMetadataUpdates.delete(videoId);
    });
  }

  // Fetch cover art eagerly at session start (after anime metadata parsing
  // settles) so new series show art on the stats timeline without requiring a
  // visit to the series detail page first.
  private prefetchCoverArtAsync(videoId: number): void {
    const pendingMetadata = this.pendingAnimeMetadataUpdates.get(videoId);
    void (async () => {
      try {
        await pendingMetadata;
        if (this.isDestroyed) {
          return;
        }
        await this.ensureCoverArt(videoId);
      } catch (error) {
        this.logger.warn('Unable to prefetch cover art', (error as Error).message);
      }
    })();
  }

  private updateVideoTitleForActiveSession(canonicalTitle: string): void {
    if (!this.sessionState) return;
    updateVideoTitleRecord(this.db, this.sessionState.videoId, canonicalTitle);
  }
}
