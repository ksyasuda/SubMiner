import crypto from 'node:crypto';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { DatabaseSync } from 'node:sqlite';
import * as fs from 'node:fs';
import { createLogger } from '../../logger';
import { pruneRetention, runRollupMaintenance } from './immersion-tracker/maintenance';
import {
  getDailyRollups,
  getMonthlyRollups,
  getQueryHints,
  getSessionSummaries,
  getSessionTimeline,
} from './immersion-tracker/query';
import {
  buildVideoKey,
  calculateTextMetrics,
  createInitialSessionState,
  deriveCanonicalTitle,
  emptyMetadata,
  hashToCode,
  isRemoteSource,
  normalizeMediaPath,
  normalizeText,
  parseFps,
  resolveBoundedInt,
  sanitizePayload,
  secToMs,
  toNullableInt,
} from './immersion-tracker/reducer';
import { enqueueWrite } from './immersion-tracker/queue';
import {
  DEFAULT_BATCH_SIZE,
  DEFAULT_DAILY_ROLLUP_RETENTION_MS,
  DEFAULT_EVENTS_RETENTION_MS,
  DEFAULT_FLUSH_INTERVAL_MS,
  DEFAULT_MAINTENANCE_INTERVAL_MS,
  DEFAULT_MAX_PAYLOAD_BYTES,
  DEFAULT_MONTHLY_ROLLUP_RETENTION_MS,
  DEFAULT_QUEUE_CAP,
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
  SCHEMA_VERSION,
  SESSION_STATUS_ACTIVE,
  SESSION_STATUS_ENDED,
  SOURCE_TYPE_LOCAL,
  SOURCE_TYPE_REMOTE,
  type ImmersionSessionRollupRow,
  type ImmersionTrackerOptions,
  type QueuedWrite,
  type SessionState,
  type SessionSummaryQueryRow,
  type SessionTimelineRow,
  type VideoMetadata,
} from './immersion-tracker/types';

export type {
  ImmersionSessionRollupRow,
  ImmersionTrackerOptions,
  ImmersionTrackerPolicy,
  SessionSummaryQueryRow,
  SessionTimelineRow,
} from './immersion-tracker/types';

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
  private readonly dailyRollupRetentionMs: number;
  private readonly monthlyRollupRetentionMs: number;
  private readonly vacuumIntervalMs: number;
  private readonly dbPath: string;
  private readonly writeLock = { locked: false };
  private flushTimer: ReturnType<typeof setTimeout> | null = null;
  private maintenanceTimer: ReturnType<typeof setInterval> | null = null;
  private flushScheduled = false;
  private droppedWriteCount = 0;
  private lastMaintenanceMs = 0;
  private lastVacuumMs = 0;
  private isDestroyed = false;
  private sessionState: SessionState | null = null;
  private currentVideoKey = '';
  private currentMediaPathOrUrl = '';
  private lastQueueWriteAtMs = 0;
  private readonly telemetryInsertStmt: ReturnType<DatabaseSync['prepare']>;
  private readonly eventInsertStmt: ReturnType<DatabaseSync['prepare']>;

  constructor(options: ImmersionTrackerOptions) {
    this.dbPath = options.dbPath;
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
    this.eventsRetentionMs =
      resolveBoundedInt(
        retention.eventsDays,
        Math.floor(DEFAULT_EVENTS_RETENTION_MS / 86_400_000),
        1,
        3650,
      ) * 86_400_000;
    this.telemetryRetentionMs =
      resolveBoundedInt(
        retention.telemetryDays,
        Math.floor(DEFAULT_TELEMETRY_RETENTION_MS / 86_400_000),
        1,
        3650,
      ) * 86_400_000;
    this.dailyRollupRetentionMs =
      resolveBoundedInt(
        retention.dailyRollupsDays,
        Math.floor(DEFAULT_DAILY_ROLLUP_RETENTION_MS / 86_400_000),
        1,
        36500,
      ) * 86_400_000;
    this.monthlyRollupRetentionMs =
      resolveBoundedInt(
        retention.monthlyRollupsDays,
        Math.floor(DEFAULT_MONTHLY_ROLLUP_RETENTION_MS / 86_400_000),
        1,
        36500,
      ) * 86_400_000;
    this.vacuumIntervalMs =
      resolveBoundedInt(
        retention.vacuumIntervalDays,
        Math.floor(DEFAULT_VACUUM_INTERVAL_MS / 86_400_000),
        1,
        3650,
      ) * 86_400_000;
    this.lastMaintenanceMs = Date.now();

    this.db = new DatabaseSync(this.dbPath);
    this.applyPragmas();
    this.ensureSchema();
    this.telemetryInsertStmt = this.db.prepare(`
      INSERT INTO imm_session_telemetry (
        session_id, sample_ms, total_watched_ms, active_watched_ms,
        lines_seen, words_seen, tokens_seen, cards_mined, lookup_count,
        lookup_hits, pause_count, pause_ms, seek_forward_count,
        seek_backward_count, media_buffer_events
      ) VALUES (
        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
      )
    `);
    this.eventInsertStmt = this.db.prepare(`
      INSERT INTO imm_session_events (
        session_id, ts_ms, event_type, line_index, segment_start_ms, segment_end_ms,
        words_delta, cards_delta, payload_json
      ) VALUES (
        ?, ?, ?, ?, ?, ?, ?, ?, ?
      )
    `);
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
    this.db.close();
  }

  async getSessionSummaries(limit = 50): Promise<SessionSummaryQueryRow[]> {
    return getSessionSummaries(this.db, limit);
  }

  async getSessionTimeline(sessionId: number, limit = 200): Promise<SessionTimelineRow[]> {
    return getSessionTimeline(this.db, sessionId, limit);
  }

  async getQueryHints(): Promise<{
    totalSessions: number;
    activeSessions: number;
  }> {
    return getQueryHints(this.db);
  }

  async getDailyRollups(limit = 60): Promise<ImmersionSessionRollupRow[]> {
    return getDailyRollups(this.db, limit);
  }

  async getMonthlyRollups(limit = 24): Promise<ImmersionSessionRollupRow[]> {
    return getMonthlyRollups(this.db, limit);
  }

  handleMediaChange(mediaPath: string | null, mediaTitle: string | null): void {
    const normalizedPath = normalizeMediaPath(mediaPath);
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
      videoId: this.getOrCreateVideo(videoKey, {
        canonicalTitle,
        sourcePath,
        sourceUrl,
        sourceType,
      }),
      startedAtMs: Date.now(),
    };

    this.logger.info(
      `Starting immersion session for path=${normalizedPath} videoId=${sessionInfo.videoId}`,
    );
    this.startSession(sessionInfo.videoId, sessionInfo.startedAtMs);
    this.captureVideoMetadataAsync(sessionInfo.videoId, sourceType, normalizedPath);
  }

  handleMediaTitleUpdate(mediaTitle: string | null): void {
    if (!this.sessionState) return;
    const normalizedTitle = normalizeText(mediaTitle);
    if (!normalizedTitle) return;
    this.currentVideoKey = normalizedTitle;
    this.updateVideoTitleForActiveSession(normalizedTitle);
  }

  recordSubtitleLine(text: string, startSec: number, endSec: number): void {
    if (!this.sessionState || !text.trim()) return;
    const cleaned = normalizeText(text);
    if (!cleaned) return;

    const metrics = calculateTextMetrics(cleaned);
    this.sessionState.currentLineIndex += 1;
    this.sessionState.linesSeen += 1;
    this.sessionState.wordsSeen += metrics.words;
    this.sessionState.tokensSeen += metrics.tokens;
    this.sessionState.pendingTelemetry = true;

    this.recordWrite({
      kind: 'event',
      sessionId: this.sessionState.sessionId,
      sampleMs: Date.now(),
      lineIndex: this.sessionState.currentLineIndex,
      segmentStartMs: secToMs(startSec),
      segmentEndMs: secToMs(endSec),
      wordsDelta: metrics.words,
      cardsDelta: 0,
      eventType: EVENT_SUBTITLE_LINE,
      payloadJson: sanitizePayload(
        {
          event: 'subtitle-line',
          text: cleaned,
          words: metrics.words,
        },
        this.maxPayloadBytes,
      ),
    });
  }

  recordPlaybackPosition(mediaTimeSec: number | null): void {
    if (!this.sessionState || mediaTimeSec === null || !Number.isFinite(mediaTimeSec)) {
      return;
    }
    const nowMs = Date.now();
    const mediaMs = Math.round(mediaTimeSec * 1000);
    if (this.sessionState.lastWallClockMs <= 0) {
      this.sessionState.lastWallClockMs = nowMs;
      this.sessionState.lastMediaMs = mediaMs;
      return;
    }

    const wallDeltaMs = nowMs - this.sessionState.lastWallClockMs;
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
            sampleMs: nowMs,
            eventType: EVENT_SEEK_FORWARD,
            wordsDelta: 0,
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
            sampleMs: nowMs,
            eventType: EVENT_SEEK_BACKWARD,
            wordsDelta: 0,
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

    this.sessionState.lastWallClockMs = nowMs;
    this.sessionState.lastMediaMs = mediaMs;
    this.sessionState.pendingTelemetry = true;
  }

  recordPauseState(isPaused: boolean): void {
    if (!this.sessionState) return;
    if (this.sessionState.isPaused === isPaused) return;

    const nowMs = Date.now();
    this.sessionState.isPaused = isPaused;
    if (isPaused) {
      this.sessionState.lastPauseStartMs = nowMs;
      this.sessionState.pauseCount += 1;
      this.recordWrite({
        kind: 'event',
        sessionId: this.sessionState.sessionId,
        sampleMs: nowMs,
        eventType: EVENT_PAUSE_START,
        cardsDelta: 0,
        wordsDelta: 0,
        payloadJson: sanitizePayload({ paused: true }, this.maxPayloadBytes),
      });
    } else {
      if (this.sessionState.lastPauseStartMs) {
        const pauseMs = Math.max(0, nowMs - this.sessionState.lastPauseStartMs);
        this.sessionState.pauseMs += pauseMs;
        this.sessionState.lastPauseStartMs = null;
      }
      this.recordWrite({
        kind: 'event',
        sessionId: this.sessionState.sessionId,
        sampleMs: nowMs,
        eventType: EVENT_PAUSE_END,
        cardsDelta: 0,
        wordsDelta: 0,
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
      sampleMs: Date.now(),
      eventType: EVENT_LOOKUP,
      cardsDelta: 0,
      wordsDelta: 0,
      payloadJson: sanitizePayload(
        {
          hit,
        },
        this.maxPayloadBytes,
      ),
    });
  }

  recordCardsMined(count = 1): void {
    if (!this.sessionState) return;
    this.sessionState.cardsMined += count;
    this.sessionState.pendingTelemetry = true;
    this.recordWrite({
      kind: 'event',
      sessionId: this.sessionState.sessionId,
      sampleMs: Date.now(),
      eventType: EVENT_CARD_MINED,
      wordsDelta: 0,
      cardsDelta: count,
      payloadJson: sanitizePayload({ cardsMined: count }, this.maxPayloadBytes),
    });
  }

  recordMediaBufferEvent(): void {
    if (!this.sessionState) return;
    this.sessionState.mediaBufferEvents += 1;
    this.sessionState.pendingTelemetry = true;
    this.recordWrite({
      kind: 'event',
      sessionId: this.sessionState.sessionId,
      sampleMs: Date.now(),
      eventType: EVENT_MEDIA_BUFFER,
      cardsDelta: 0,
      wordsDelta: 0,
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
    this.lastQueueWriteAtMs = Date.now();
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
      sampleMs: Date.now(),
      totalWatchedMs: this.sessionState.totalWatchedMs,
      activeWatchedMs: this.sessionState.activeWatchedMs,
      linesSeen: this.sessionState.linesSeen,
      wordsSeen: this.sessionState.wordsSeen,
      tokensSeen: this.sessionState.tokensSeen,
      cardsMined: this.sessionState.cardsMined,
      lookupCount: this.sessionState.lookupCount,
      lookupHits: this.sessionState.lookupHits,
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

  private flushSingle(write: QueuedWrite): void {
    if (write.kind === 'telemetry') {
      this.telemetryInsertStmt.run(
        write.sessionId,
        write.sampleMs!,
        write.totalWatchedMs!,
        write.activeWatchedMs!,
        write.linesSeen!,
        write.wordsSeen!,
        write.tokensSeen!,
        write.cardsMined!,
        write.lookupCount!,
        write.lookupHits!,
        write.pauseCount!,
        write.pauseMs!,
        write.seekForwardCount!,
        write.seekBackwardCount!,
        write.mediaBufferEvents!,
      );
      return;
    }

    this.eventInsertStmt.run(
      write.sessionId,
      write.sampleMs!,
      write.eventType!,
      write.lineIndex ?? null,
      write.segmentStartMs ?? null,
      write.segmentEndMs ?? null,
      write.wordsDelta ?? 0,
      write.cardsDelta ?? 0,
      write.payloadJson ?? null,
    );
  }

  private applyPragmas(): void {
    this.db.exec('PRAGMA journal_mode = WAL');
    this.db.exec('PRAGMA synchronous = NORMAL');
    this.db.exec('PRAGMA foreign_keys = ON');
    this.db.exec('PRAGMA busy_timeout = 2500');
  }

  private ensureSchema(): void {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS imm_schema_version (
        schema_version INTEGER PRIMARY KEY,
        applied_at_ms INTEGER NOT NULL
      );
    `);

    const currentVersion = this.db
      .prepare('SELECT schema_version FROM imm_schema_version ORDER BY schema_version DESC LIMIT 1')
      .get() as { schema_version: number } | null;
    if (currentVersion?.schema_version === SCHEMA_VERSION) {
      return;
    }

    this.db.exec(`
      CREATE TABLE IF NOT EXISTS imm_videos(
        video_id INTEGER PRIMARY KEY AUTOINCREMENT,
        video_key TEXT NOT NULL UNIQUE,
        canonical_title TEXT NOT NULL,
        source_type INTEGER NOT NULL,
        source_path TEXT,
        source_url TEXT,
        duration_ms INTEGER NOT NULL CHECK(duration_ms>=0),
        file_size_bytes INTEGER CHECK(file_size_bytes>=0),
        codec_id INTEGER, container_id INTEGER,
        width_px INTEGER, height_px INTEGER, fps_x100 INTEGER,
        bitrate_kbps INTEGER, audio_codec_id INTEGER,
        hash_sha256 TEXT, screenshot_path TEXT,
        metadata_json TEXT,
        created_at_ms INTEGER NOT NULL, updated_at_ms INTEGER NOT NULL
      );
    `);
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS imm_sessions(
        session_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_uuid TEXT NOT NULL UNIQUE,
        video_id INTEGER NOT NULL,
        started_at_ms INTEGER NOT NULL, ended_at_ms INTEGER,
        status INTEGER NOT NULL,
        locale_id INTEGER, target_lang_id INTEGER,
        difficulty_tier INTEGER, subtitle_mode INTEGER,
        created_at_ms INTEGER NOT NULL, updated_at_ms INTEGER NOT NULL,
        FOREIGN KEY(video_id) REFERENCES imm_videos(video_id)
      );
    `);
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS imm_session_telemetry(
        telemetry_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        sample_ms INTEGER NOT NULL,
        total_watched_ms INTEGER NOT NULL DEFAULT 0,
        active_watched_ms INTEGER NOT NULL DEFAULT 0,
        lines_seen INTEGER NOT NULL DEFAULT 0,
        words_seen INTEGER NOT NULL DEFAULT 0,
        tokens_seen INTEGER NOT NULL DEFAULT 0,
        cards_mined INTEGER NOT NULL DEFAULT 0,
        lookup_count INTEGER NOT NULL DEFAULT 0,
        lookup_hits INTEGER NOT NULL DEFAULT 0,
        pause_count INTEGER NOT NULL DEFAULT 0,
        pause_ms INTEGER NOT NULL DEFAULT 0,
        seek_forward_count INTEGER NOT NULL DEFAULT 0,
        seek_backward_count INTEGER NOT NULL DEFAULT 0,
        media_buffer_events INTEGER NOT NULL DEFAULT 0,
        FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
      );
    `);
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS imm_session_events(
        event_id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        ts_ms INTEGER NOT NULL,
        event_type INTEGER NOT NULL,
        line_index INTEGER,
        segment_start_ms INTEGER,
        segment_end_ms INTEGER,
        words_delta INTEGER NOT NULL DEFAULT 0,
        cards_delta INTEGER NOT NULL DEFAULT 0,
        payload_json TEXT,
        FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
      );
    `);
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS imm_daily_rollups(
        rollup_day INTEGER NOT NULL,
        video_id INTEGER,
        total_sessions INTEGER NOT NULL DEFAULT 0,
        total_active_min REAL NOT NULL DEFAULT 0,
        total_lines_seen INTEGER NOT NULL DEFAULT 0,
        total_words_seen INTEGER NOT NULL DEFAULT 0,
        total_tokens_seen INTEGER NOT NULL DEFAULT 0,
        total_cards INTEGER NOT NULL DEFAULT 0,
        cards_per_hour REAL,
        words_per_min REAL,
        lookup_hit_rate REAL,
        PRIMARY KEY (rollup_day, video_id)
      );
    `);
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS imm_monthly_rollups(
        rollup_month INTEGER NOT NULL,
        video_id INTEGER,
        total_sessions INTEGER NOT NULL DEFAULT 0,
        total_active_min REAL NOT NULL DEFAULT 0,
        total_lines_seen INTEGER NOT NULL DEFAULT 0,
        total_words_seen INTEGER NOT NULL DEFAULT 0,
        total_tokens_seen INTEGER NOT NULL DEFAULT 0,
        total_cards INTEGER NOT NULL DEFAULT 0,
        PRIMARY KEY (rollup_month, video_id)
      );
    `);

    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_sessions_video_started
      ON imm_sessions(video_id, started_at_ms DESC)
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_sessions_status_started
      ON imm_sessions(status, started_at_ms DESC)
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_telemetry_session_sample
      ON imm_session_telemetry(session_id, sample_ms DESC)
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_events_session_ts
      ON imm_session_events(session_id, ts_ms DESC)
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_events_type_ts
      ON imm_session_events(event_type, ts_ms DESC)
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_rollups_day_video
      ON imm_daily_rollups(rollup_day, video_id)
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_rollups_month_video
      ON imm_monthly_rollups(rollup_month, video_id)
    `);

    this.db.exec(`
      INSERT INTO imm_schema_version(schema_version, applied_at_ms)
      VALUES (${SCHEMA_VERSION}, ${Date.now()})
      ON CONFLICT DO NOTHING
    `);
  }

  private scheduleMaintenance(): void {
    this.maintenanceTimer = setInterval(() => {
      this.runMaintenance();
    }, this.maintenanceIntervalMs);
    this.runMaintenance();
  }

  private runMaintenance(): void {
    if (this.isDestroyed) return;
    try {
      this.flushTelemetry(true);
      this.flushNow();
      const nowMs = Date.now();
      pruneRetention(this.db, nowMs, {
        eventsRetentionMs: this.eventsRetentionMs,
        telemetryRetentionMs: this.telemetryRetentionMs,
        dailyRollupRetentionMs: this.dailyRollupRetentionMs,
        monthlyRollupRetentionMs: this.monthlyRollupRetentionMs,
      });
      runRollupMaintenance(this.db);

      if (nowMs - this.lastVacuumMs >= this.vacuumIntervalMs && !this.writeLock.locked) {
        this.db.exec('VACUUM');
        this.lastVacuumMs = nowMs;
      }
      this.lastMaintenanceMs = nowMs;
    } catch (error) {
      this.logger.warn(
        'Immersion tracker maintenance failed, will retry later',
        (error as Error).message,
      );
    }
  }

  private runRollupMaintenance(): void {
    runRollupMaintenance(this.db);
  }

  private startSession(videoId: number, startedAtMs?: number): void {
    const nowMs = startedAtMs ?? Date.now();
    const result = this.startSessionStatement(videoId, nowMs);
    const sessionId = Number(result.lastInsertRowid);
    this.sessionState = createInitialSessionState(sessionId, videoId, nowMs);
    this.recordWrite({
      kind: 'telemetry',
      sessionId,
      sampleMs: nowMs,
      totalWatchedMs: 0,
      activeWatchedMs: 0,
      linesSeen: 0,
      wordsSeen: 0,
      tokensSeen: 0,
      cardsMined: 0,
      lookupCount: 0,
      lookupHits: 0,
      pauseCount: 0,
      pauseMs: 0,
      seekForwardCount: 0,
      seekBackwardCount: 0,
      mediaBufferEvents: 0,
    });
    this.scheduleFlush(0);
  }

  private startSessionStatement(
    videoId: number,
    startedAtMs: number,
  ): {
    lastInsertRowid: number | bigint;
  } {
    const sessionUuid = crypto.randomUUID();
    return this.db
      .prepare(
        `
        INSERT INTO imm_sessions (
          session_uuid, video_id, started_at_ms, status, created_at_ms, updated_at_ms
        ) VALUES (?, ?, ?, ?, ?, ?)
      `,
      )
      .run(sessionUuid, videoId, startedAtMs, SESSION_STATUS_ACTIVE, startedAtMs, startedAtMs);
  }

  private finalizeActiveSession(): void {
    if (!this.sessionState) return;
    const endedAt = Date.now();
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

    this.db
      .prepare(
        'UPDATE imm_sessions SET ended_at_ms = ?, status = ?, updated_at_ms = ? WHERE session_id = ?',
      )
      .run(endedAt, SESSION_STATUS_ENDED, Date.now(), this.sessionState.sessionId);
    this.sessionState = null;
  }

  private getOrCreateVideo(
    videoKey: string,
    details: {
      canonicalTitle: string;
      sourcePath: string | null;
      sourceUrl: string | null;
      sourceType: number;
    },
  ): number {
    const existing = this.db
      .prepare('SELECT video_id FROM imm_videos WHERE video_key = ?')
      .get(videoKey) as { video_id: number } | null;
    if (existing?.video_id) {
      this.db
        .prepare('UPDATE imm_videos SET canonical_title = ?, updated_at_ms = ? WHERE video_id = ?')
        .run(details.canonicalTitle || 'unknown', Date.now(), existing.video_id);
      return existing.video_id;
    }

    const nowMs = Date.now();
    const insert = this.db.prepare(`
      INSERT INTO imm_videos (
        video_key, canonical_title, source_type, source_path, source_url,
        duration_ms, file_size_bytes, codec_id, container_id, width_px, height_px,
        fps_x100, bitrate_kbps, audio_codec_id, hash_sha256, screenshot_path,
        metadata_json, created_at_ms, updated_at_ms
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const result = insert.run(
      videoKey,
      details.canonicalTitle || 'unknown',
      details.sourceType,
      details.sourcePath,
      details.sourceUrl,
      0,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      nowMs,
      nowMs,
    );
    return Number(result.lastInsertRowid);
  }

  private updateVideoMetadata(videoId: number, metadata: VideoMetadata): void {
    this.db
      .prepare(
        `
        UPDATE imm_videos
        SET
          duration_ms = ?,
          file_size_bytes = ?,
          codec_id = ?,
          container_id = ?,
          width_px = ?,
          height_px = ?,
          fps_x100 = ?,
          bitrate_kbps = ?,
          audio_codec_id = ?,
          hash_sha256 = ?,
          screenshot_path = ?,
          metadata_json = ?,
          updated_at_ms = ?
        WHERE video_id = ?
      `,
      )
      .run(
        metadata.durationMs,
        metadata.fileSizeBytes,
        metadata.codecId,
        metadata.containerId,
        metadata.widthPx,
        metadata.heightPx,
        metadata.fpsX100,
        metadata.bitrateKbps,
        metadata.audioCodecId,
        metadata.hashSha256,
        metadata.screenshotPath,
        metadata.metadataJson,
        Date.now(),
        videoId,
      );
  }

  private captureVideoMetadataAsync(videoId: number, sourceType: number, mediaPath: string): void {
    if (sourceType !== SOURCE_TYPE_LOCAL) return;
    void (async () => {
      try {
        const metadata = await this.getLocalVideoMetadata(mediaPath);
        this.updateVideoMetadata(videoId, metadata);
      } catch (error) {
        this.logger.warn('Unable to capture local video metadata', (error as Error).message);
      }
    })();
  }

  private async getLocalVideoMetadata(mediaPath: string): Promise<VideoMetadata> {
    const hash = await this.computeSha256(mediaPath);
    const info = await this.runFfprobe(mediaPath);
    const stat = await fs.promises.stat(mediaPath);
    return {
      sourceType: SOURCE_TYPE_LOCAL,
      canonicalTitle: deriveCanonicalTitle(mediaPath),
      durationMs: info.durationMs || 0,
      fileSizeBytes: Number.isFinite(stat.size) ? stat.size : null,
      codecId: info.codecId ?? null,
      containerId: info.containerId ?? null,
      widthPx: info.widthPx ?? null,
      heightPx: info.heightPx ?? null,
      fpsX100: info.fpsX100 ?? null,
      bitrateKbps: info.bitrateKbps ?? null,
      audioCodecId: info.audioCodecId ?? null,
      hashSha256: hash,
      screenshotPath: null,
      metadataJson: null,
    };
  }

  private async computeSha256(mediaPath: string): Promise<string | null> {
    return new Promise((resolve) => {
      const file = fs.createReadStream(mediaPath);
      const digest = crypto.createHash('sha256');
      file.on('data', (chunk) => digest.update(chunk));
      file.on('end', () => resolve(digest.digest('hex')));
      file.on('error', () => resolve(null));
    });
  }

  private runFfprobe(mediaPath: string): Promise<{
    durationMs: number | null;
    codecId: number | null;
    containerId: number | null;
    widthPx: number | null;
    heightPx: number | null;
    fpsX100: number | null;
    bitrateKbps: number | null;
    audioCodecId: number | null;
  }> {
    return new Promise((resolve) => {
      const child = spawn('ffprobe', [
        '-v',
        'error',
        '-print_format',
        'json',
        '-show_entries',
        'stream=codec_type,codec_tag_string,width,height,avg_frame_rate,bit_rate',
        '-show_entries',
        'format=duration,bit_rate',
        mediaPath,
      ]);

      let output = '';
      let errorOutput = '';
      child.stdout.on('data', (chunk) => {
        output += chunk.toString('utf-8');
      });
      child.stderr.on('data', (chunk) => {
        errorOutput += chunk.toString('utf-8');
      });
      child.on('error', () => resolve(emptyMetadata()));
      child.on('close', () => {
        if (errorOutput && output.length === 0) {
          resolve(emptyMetadata());
          return;
        }

        try {
          const parsed = JSON.parse(output) as {
            format?: { duration?: string; bit_rate?: string };
            streams?: Array<{
              codec_type?: string;
              codec_tag_string?: string;
              width?: number;
              height?: number;
              avg_frame_rate?: string;
              bit_rate?: string;
            }>;
          };

          const durationText = parsed.format?.duration;
          const bitrateText = parsed.format?.bit_rate;
          const durationMs = Number(durationText) ? Math.round(Number(durationText) * 1000) : null;
          const bitrateKbps = Number(bitrateText) ? Math.round(Number(bitrateText) / 1000) : null;

          let codecId: number | null = null;
          let containerId: number | null = null;
          let widthPx: number | null = null;
          let heightPx: number | null = null;
          let fpsX100: number | null = null;
          let audioCodecId: number | null = null;

          for (const stream of parsed.streams ?? []) {
            if (stream.codec_type === 'video') {
              widthPx = toNullableInt(stream.width);
              heightPx = toNullableInt(stream.height);
              fpsX100 = parseFps(stream.avg_frame_rate);
              codecId = hashToCode(stream.codec_tag_string);
              containerId = 0;
            }
            if (stream.codec_type === 'audio') {
              audioCodecId = hashToCode(stream.codec_tag_string);
              if (audioCodecId && audioCodecId > 0) {
                break;
              }
            }
          }

          resolve({
            durationMs,
            codecId,
            containerId,
            widthPx,
            heightPx,
            fpsX100,
            bitrateKbps,
            audioCodecId,
          });
        } catch {
          resolve(emptyMetadata());
        }
      });
    });
  }

  private updateVideoTitleForActiveSession(canonicalTitle: string): void {
    if (!this.sessionState) return;
    this.db
      .prepare('UPDATE imm_videos SET canonical_title = ?, updated_at_ms = ? WHERE video_id = ?')
      .run(canonicalTitle, Date.now(), this.sessionState.videoId);
  }
}
