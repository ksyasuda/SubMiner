export const SCHEMA_VERSION = 9;
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
  resolveLegacyVocabularyPos?: (
    row: LegacyVocabularyPosRow,
  ) => Promise<LegacyVocabularyPosResolution | null>;
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
  markedWatched: boolean;
}

interface QueuedTelemetryWrite {
  kind: 'telemetry';
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

interface QueuedEventWrite {
  kind: 'event';
  sessionId: number;
  sampleMs?: number;
  eventType?: number;
  lineIndex?: number | null;
  segmentStartMs?: number | null;
  segmentEndMs?: number | null;
  wordsDelta?: number;
  cardsDelta?: number;
  payloadJson?: string | null;
}

interface QueuedWordWrite {
  kind: 'word';
  headword: string;
  word: string;
  reading: string;
  partOfSpeech: string;
  pos1: string;
  pos2: string;
  pos3: string;
  firstSeen: number;
  lastSeen: number;
  frequencyRank: number | null;
}

interface QueuedKanjiWrite {
  kind: 'kanji';
  kanji: string;
  firstSeen: number;
  lastSeen: number;
}

export interface CountedWordOccurrence {
  headword: string;
  word: string;
  reading: string;
  partOfSpeech: string;
  pos1: string;
  pos2: string;
  pos3: string;
  occurrenceCount: number;
  frequencyRank: number | null;
}

export interface CountedKanjiOccurrence {
  kanji: string;
  occurrenceCount: number;
}

interface QueuedSubtitleLineWrite {
  kind: 'subtitleLine';
  sessionId: number;
  videoId: number;
  lineIndex: number;
  segmentStartMs: number | null;
  segmentEndMs: number | null;
  text: string;
  wordOccurrences: CountedWordOccurrence[];
  kanjiOccurrences: CountedKanjiOccurrence[];
  firstSeen: number;
  lastSeen: number;
}

export type QueuedWrite =
  | QueuedTelemetryWrite
  | QueuedEventWrite
  | QueuedWordWrite
  | QueuedKanjiWrite
  | QueuedSubtitleLineWrite;

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

export interface ParsedAnimeVideoMetadata {
  animeId: number | null;
  parsedBasename: string | null;
  parsedTitle: string | null;
  parsedSeason: number | null;
  parsedEpisode: number | null;
  parserSource: string | null;
  parserConfidence: number | null;
  parseMetadataJson: string | null;
}

export interface ParsedAnimeVideoGuess {
  parsedBasename: string | null;
  parsedTitle: string;
  parsedSeason: number | null;
  parsedEpisode: number | null;
  parserSource: 'guessit' | 'fallback';
  parserConfidence: number;
  parseMetadataJson: string;
}

export interface SessionSummaryQueryRow {
  sessionId: number;
  videoId: number | null;
  canonicalTitle: string | null;
  animeId: number | null;
  animeTitle: string | null;
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

export interface VocabularyStatsRow {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  partOfSpeech: string | null;
  pos1: string | null;
  pos2: string | null;
  pos3: string | null;
  frequency: number;
  frequencyRank: number | null;
  firstSeen: number;
  lastSeen: number;
}

export interface VocabularyCleanupSummary {
  scanned: number;
  kept: number;
  deleted: number;
  repaired: number;
}

export interface LegacyVocabularyPosRow {
  headword: string;
  word: string;
  reading: string | null;
}

export interface LegacyVocabularyPosResolution {
  headword: string;
  reading: string;
  partOfSpeech: string;
  pos1: string;
  pos2: string;
  pos3: string;
}

export interface KanjiStatsRow {
  kanjiId: number;
  kanji: string;
  frequency: number;
  firstSeen: number;
  lastSeen: number;
}

export interface WordOccurrenceRow {
  animeId: number | null;
  animeTitle: string | null;
  videoId: number;
  videoTitle: string;
  sessionId: number;
  lineIndex: number;
  segmentStartMs: number | null;
  segmentEndMs: number | null;
  text: string;
  occurrenceCount: number;
}

export interface KanjiOccurrenceRow {
  animeId: number | null;
  animeTitle: string | null;
  videoId: number;
  videoTitle: string;
  sessionId: number;
  lineIndex: number;
  segmentStartMs: number | null;
  segmentEndMs: number | null;
  text: string;
  occurrenceCount: number;
}

export interface SessionEventRow {
  eventType: number;
  tsMs: number;
  payload: string | null;
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

export interface MediaArtRow {
  videoId: number;
  anilistId: number | null;
  coverUrl: string | null;
  coverBlob: Buffer | null;
  titleRomaji: string | null;
  titleEnglish: string | null;
  episodesTotal: number | null;
  fetchedAtMs: number;
}

export interface MediaLibraryRow {
  videoId: number;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  lastWatchedMs: number;
  hasCoverArt: number;
}

export interface MediaDetailRow {
  videoId: number;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  totalLinesSeen: number;
  totalLookupCount: number;
  totalLookupHits: number;
}

export interface AnimeLibraryRow {
  animeId: number;
  canonicalTitle: string;
  anilistId: number | null;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  episodeCount: number;
  episodesTotal: number | null;
  lastWatchedMs: number;
}

export interface AnimeDetailRow {
  animeId: number;
  canonicalTitle: string;
  anilistId: number | null;
  titleRomaji: string | null;
  titleEnglish: string | null;
  titleNative: string | null;
  description: string | null;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  totalLinesSeen: number;
  totalLookupCount: number;
  totalLookupHits: number;
  episodeCount: number;
  lastWatchedMs: number;
}

export interface AnimeAnilistEntryRow {
  anilistId: number;
  titleRomaji: string | null;
  titleEnglish: string | null;
  season: number | null;
}

export interface AnimeEpisodeRow {
  animeId: number;
  videoId: number;
  canonicalTitle: string;
  parsedTitle: string | null;
  season: number | null;
  episode: number | null;
  durationMs: number;
  watched: number;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  lastWatchedMs: number;
}

export interface StreakCalendarRow {
  epochDay: number;
  totalActiveMin: number;
}

export interface AnimeWordRow {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  partOfSpeech: string | null;
  frequency: number;
}

export interface EpisodesPerDayRow {
  epochDay: number;
  episodeCount: number;
}

export interface NewAnimePerDayRow {
  epochDay: number;
  newAnimeCount: number;
}

export interface WatchTimePerAnimeRow {
  epochDay: number;
  animeId: number;
  animeTitle: string;
  totalActiveMin: number;
}

export interface WordDetailRow {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  partOfSpeech: string | null;
  pos1: string | null;
  pos2: string | null;
  pos3: string | null;
  frequency: number;
  firstSeen: number;
  lastSeen: number;
}

export interface WordAnimeAppearanceRow {
  animeId: number;
  animeTitle: string;
  occurrenceCount: number;
}

export interface SimilarWordRow {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  frequency: number;
}

export interface KanjiDetailRow {
  kanjiId: number;
  kanji: string;
  frequency: number;
  firstSeen: number;
  lastSeen: number;
}

export interface KanjiAnimeAppearanceRow {
  animeId: number;
  animeTitle: string;
  occurrenceCount: number;
}

export interface KanjiWordRow {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  frequency: number;
}

export interface EpisodeCardEventRow {
  eventId: number;
  sessionId: number;
  tsMs: number;
  cardsDelta: number;
  noteIds: number[];
}
