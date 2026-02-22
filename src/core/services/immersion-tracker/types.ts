export const SCHEMA_VERSION = 1;
export const DEFAULT_QUEUE_CAP = 1_000;
export const DEFAULT_BATCH_SIZE = 25;
export const DEFAULT_FLUSH_INTERVAL_MS = 500;
export const DEFAULT_MAINTENANCE_INTERVAL_MS = 24 * 60 * 60 * 1000;
const ONE_WEEK_MS = 7 * 24 * 60 * 60 * 1000;
export const DEFAULT_EVENTS_RETENTION_MS = ONE_WEEK_MS;
export const DEFAULT_VACUUM_INTERVAL_MS = ONE_WEEK_MS;
export const DEFAULT_TELEMETRY_RETENTION_MS = 30 * 24 * 60 * 60 * 1000;
export const DEFAULT_DAILY_ROLLUP_RETENTION_MS = 365 * 24 * 60 * 60 * 1000;
export const DEFAULT_MONTHLY_ROLLUP_RETENTION_MS = 5 * 365 * 24 * 60 * 60 * 1000;
export const DEFAULT_MAX_PAYLOAD_BYTES = 256;

export const SOURCE_TYPE_LOCAL = 1;
export const SOURCE_TYPE_REMOTE = 2;

export const SESSION_STATUS_ACTIVE = 1;
export const SESSION_STATUS_ENDED = 2;

export const EVENT_SUBTITLE_LINE = 1;
export const EVENT_MEDIA_BUFFER = 2;
export const EVENT_LOOKUP = 3;
export const EVENT_CARD_MINED = 4;
export const EVENT_SEEK_FORWARD = 5;
export const EVENT_SEEK_BACKWARD = 6;
export const EVENT_PAUSE_START = 7;
export const EVENT_PAUSE_END = 8;

export interface ImmersionTrackerOptions {
  dbPath: string;
  policy?: ImmersionTrackerPolicy;
}

export interface ImmersionTrackerPolicy {
  queueCap?: number;
  batchSize?: number;
  flushIntervalMs?: number;
  maintenanceIntervalMs?: number;
  payloadCapBytes?: number;
  retention?: {
    eventsDays?: number;
    telemetryDays?: number;
    dailyRollupsDays?: number;
    monthlyRollupsDays?: number;
    vacuumIntervalDays?: number;
  };
}

export interface TelemetryAccumulator {
  totalWatchedMs: number;
  activeWatchedMs: number;
  linesSeen: number;
  wordsSeen: number;
  tokensSeen: number;
  cardsMined: number;
  lookupCount: number;
  lookupHits: number;
  pauseCount: number;
  pauseMs: number;
  seekForwardCount: number;
  seekBackwardCount: number;
  mediaBufferEvents: number;
}

export interface SessionState extends TelemetryAccumulator {
  sessionId: number;
  videoId: number;
  startedAtMs: number;
  currentLineIndex: number;
  lastWallClockMs: number;
  lastMediaMs: number | null;
  lastPauseStartMs: number | null;
  isPaused: boolean;
  pendingTelemetry: boolean;
}

export interface QueuedWrite {
  kind: 'telemetry' | 'event';
  sessionId: number;
  sampleMs?: number;
  totalWatchedMs?: number;
  activeWatchedMs?: number;
  linesSeen?: number;
  wordsSeen?: number;
  tokensSeen?: number;
  cardsMined?: number;
  lookupCount?: number;
  lookupHits?: number;
  pauseCount?: number;
  pauseMs?: number;
  seekForwardCount?: number;
  seekBackwardCount?: number;
  mediaBufferEvents?: number;
  eventType?: number;
  lineIndex?: number | null;
  segmentStartMs?: number | null;
  segmentEndMs?: number | null;
  wordsDelta?: number;
  cardsDelta?: number;
  payloadJson?: string | null;
}

export interface VideoMetadata {
  sourceType: number;
  canonicalTitle: string;
  durationMs: number;
  fileSizeBytes: number | null;
  codecId: number | null;
  containerId: number | null;
  widthPx: number | null;
  heightPx: number | null;
  fpsX100: number | null;
  bitrateKbps: number | null;
  audioCodecId: number | null;
  hashSha256: string | null;
  screenshotPath: string | null;
  metadataJson: string | null;
}

export interface SessionSummaryQueryRow {
  videoId: number | null;
  startedAtMs: number;
  endedAtMs: number | null;
  totalWatchedMs: number;
  activeWatchedMs: number;
  linesSeen: number;
  wordsSeen: number;
  tokensSeen: number;
  cardsMined: number;
  lookupCount: number;
  lookupHits: number;
}

export interface SessionTimelineRow {
  sampleMs: number;
  totalWatchedMs: number;
  activeWatchedMs: number;
  linesSeen: number;
  wordsSeen: number;
  tokensSeen: number;
  cardsMined: number;
}

export interface ImmersionSessionRollupRow {
  rollupDayOrMonth: number;
  videoId: number | null;
  totalSessions: number;
  totalActiveMin: number;
  totalLinesSeen: number;
  totalWordsSeen: number;
  totalTokensSeen: number;
  totalCards: number;
  cardsPerHour: number | null;
  wordsPerMin: number | null;
  lookupHitRate: number | null;
}

export interface ProbeMetadata {
  durationMs: number | null;
  codecId: number | null;
  containerId: number | null;
  widthPx: number | null;
  heightPx: number | null;
  fpsX100: number | null;
  bitrateKbps: number | null;
  audioCodecId: number | null;
}
