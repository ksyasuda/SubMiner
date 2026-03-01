import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import * as fs from 'node:fs';
import { createLogger } from '../../logger';
import { getLocalVideoMetadata } from './immersion-tracker/metadata';
import { pruneRetention, runRollupMaintenance } from './immersion-tracker/maintenance';
import { finalizeSessionRecord, startSessionRecord } from './immersion-tracker/session';
import {
  applyPragmas,
  createTrackerPreparedStatements,
  ensureSchema,
  executeQueuedWrite,
  getOrCreateVideoRecord,
  type TrackerPreparedStatements,
  updateVideoMetadataRecord,
  updateVideoTitleRecord,
} from './immersion-tracker/storage';
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
  extractLineVocabulary,
  deriveCanonicalTitle,
  isRemoteSource,
  normalizeMediaPath,
  normalizeText,
  resolveBoundedInt,
  sanitizePayload,
  secToMs,
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
  SOURCE_TYPE_LOCAL,
  SOURCE_TYPE_REMOTE,
  type ImmersionSessionRollupRow,
  type ImmersionTrackerOptions,
  type QueuedWrite,
  type SessionState,
  type SessionSummaryQueryRow,
  type SessionTimelineRow,
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
  private lastVacuumMs = 0;
  private isDestroyed = false;
  private sessionState: SessionState | null = null;
  private currentVideoKey = '';
  private currentMediaPathOrUrl = '';
  private readonly preparedStatements: TrackerPreparedStatements;

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
    this.db = new DatabaseSync(this.dbPath);
    applyPragmas(this.db);
    ensureSchema(this.db);
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
      videoId: getOrCreateVideoRecord(this.db, videoKey, {
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
    const nowMs = Date.now();
    const nowSec = nowMs / 1000;

    const metrics = calculateTextMetrics(cleaned);
    const extractedVocabulary = extractLineVocabulary(cleaned);
    this.sessionState.currentLineIndex += 1;
    this.sessionState.linesSeen += 1;
    this.sessionState.wordsSeen += metrics.words;
    this.sessionState.tokensSeen += metrics.tokens;
    this.sessionState.pendingTelemetry = true;

    for (const { headword, word, reading } of extractedVocabulary.words) {
      this.recordWrite({
        kind: 'word',
        headword,
        word,
        reading,
        firstSeen: nowSec,
        lastSeen: nowSec,
      });
    }

    for (const kanji of extractedVocabulary.kanji) {
      this.recordWrite({
        kind: 'kanji',
        kanji,
        firstSeen: nowSec,
        lastSeen: nowSec,
      });
    }

    this.recordWrite({
      kind: 'event',
      sessionId: this.sessionState.sessionId,
      sampleMs: nowMs,
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
    executeQueuedWrite(write, this.preparedStatements);
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
      this.runRollupMaintenance();

      if (nowMs - this.lastVacuumMs >= this.vacuumIntervalMs && !this.writeLock.locked) {
        this.db.exec('VACUUM');
        this.lastVacuumMs = nowMs;
      }
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
    const { sessionId, state } = startSessionRecord(this.db, videoId, startedAtMs);
    this.sessionState = state;
    this.recordWrite({
      kind: 'telemetry',
      sessionId,
      sampleMs: state.startedAtMs,
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

    finalizeSessionRecord(this.db, this.sessionState, endedAt);
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

  private updateVideoTitleForActiveSession(canonicalTitle: string): void {
    if (!this.sessionState) return;
    updateVideoTitleRecord(this.db, this.sessionState.videoId, canonicalTitle);
  }
}
