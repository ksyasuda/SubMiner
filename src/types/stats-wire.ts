import type { MediaKind, TmdbMediaType } from '../shared/media-kind';

export interface SessionSummary {
  sessionId: number;
  canonicalTitle: string | null;
  videoId: number | null;
  animeId: number | null;
  animeTitle: string | null;
  startedAtMs: number;
  endedAtMs: number | null;
  totalWatchedMs: number;
  activeWatchedMs: number;
  linesSeen: number;
  tokensSeen: number;
  cardsMined: number;
  lookupCount: number;
  lookupHits: number;
  yomitanLookupCount: number;
  knownWordsSeen: number;
  knownWordRate: number;
}

export interface DailyRollup {
  rollupDayOrMonth: number;
  videoId: number | null;
  totalSessions: number;
  totalActiveMin: number;
  totalLinesSeen: number;
  totalTokensSeen: number;
  totalCards: number;
  cardsPerHour: number | null;
  tokensPerMin: number | null;
  lookupHitRate: number | null;
}

export type MonthlyRollup = DailyRollup;

export interface SessionTimelinePoint {
  sampleMs: number;
  totalWatchedMs: number;
  activeWatchedMs: number;
  linesSeen: number;
  tokensSeen: number;
  cardsMined: number;
}

export interface SessionEvent {
  eventType: number;
  tsMs: number;
  payload: string | null;
}

export interface AnkiNotePreview {
  word: string;
  sentence: string;
  translation: string;
}

export interface StatsAnkiNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
  preview?: AnkiNotePreview;
}

export interface VocabularyEntry {
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
  animeCount: number;
  firstSeen: number;
  lastSeen: number;
}

/** One written form of a dictionary word, such as 見て for 見る. */
export interface DictionaryWordForm {
  wordId: number;
  word: string;
  reading: string;
  seen: number;
}

/** A dictionary word with every form it was seen in folded together. */
export interface DictionaryWordRow {
  /** The form to open: the dictionary form when it was seen, otherwise the most seen one. */
  wordId: number;
  headword: string;
  /** Reading of the dictionary form, or empty when only conjugated forms were seen. */
  reading: string;
  partOfSpeech: string | null;
  rank: number | null;
  seen: number;
  /** Titles it appeared in. Computed for the returned rows. */
  titles: number;
  firstSeen: number;
  lastSeen: number;
  known: boolean;
  isName: boolean;
  /** Every line with this word came from a stream, so it can't be mined. */
  streamOnly: boolean;
  formCount: number;
  /** The most seen forms, capped so a word with hundreds of forms stays small. */
  forms: DictionaryWordForm[];
}

export interface DictionaryWordPage {
  rows: DictionaryWordRow[];
  /** Words matching the query across all pages. */
  total: number;
  /** Highest `seen` across all pages, for scaling inline bars. */
  maxSeen: number;
}

export const VOCABULARY_COVERAGE_BANDS = ['top1k', 'top5k', 'top10k', 'top20k'] as const;
export type VocabularyCoverageBandId = (typeof VOCABULARY_COVERAGE_BANDS)[number];

/** Words seen in a frequency band, and how many of those are known. Names are left out. */
export interface VocabularyCoverageBand {
  band: VocabularyCoverageBandId;
  seen: number;
  known: number;
}

export interface StatsExcludedWord {
  headword: string;
  word: string;
  reading: string;
}

/** One animation burst the duplicate-line cleanup found in the stats database. */
export interface StatsDuplicateLineSample {
  videoId: number;
  videoTitle: string | null;
  text: string;
  /** Events recorded for this run, including the one that is kept. */
  frames: number;
  removedLines: number;
  startMs: number;
  endMs: number;
}

export interface StatsDuplicateLineCleanupResult {
  dryRun: boolean;
  lookbackDays: number | null;
  scannedLines: number;
  burstGroups: number;
  removedLines: number;
  removedWordOccurrences: number;
  removedKanjiOccurrences: number;
  samples: StatsDuplicateLineSample[];
}

export interface StatsCoverImage {
  contentType: string;
  dataUrl: string;
}

export interface StatsCoverImagesData {
  anime: Record<number, StatsCoverImage | null>;
  media: Record<number, StatsCoverImage | null>;
}

export interface KanjiEntry {
  kanjiId: number;
  kanji: string;
  frequency: number;
  firstSeen: number;
  lastSeen: number;
}

export interface VocabularyOccurrenceEntry {
  animeId: number | null;
  animeTitle: string | null;
  videoId: number;
  videoTitle: string;
  sourcePath: string | null;
  secondaryText: string | null;
  sessionId: number;
  lineIndex: number;
  segmentStartMs: number | null;
  segmentEndMs: number | null;
  text: string;
  occurrenceCount: number;
}

export interface SentenceSearchResult {
  animeId: number | null;
  animeTitle: string | null;
  videoId: number;
  videoTitle: string;
  sourcePath: string | null;
  secondaryText: string | null;
  sessionId: number;
  lineIndex: number;
  segmentStartMs: number | null;
  segmentEndMs: number | null;
  text: string;
}

export interface OverviewData {
  sessions: SessionSummary[];
  rollups: DailyRollup[];
  hints: {
    totalSessions: number;
    activeSessions: number;
    episodesToday: number;
    activeAnimeCount: number;
    totalEpisodesWatched: number;
    totalAnimeCompleted: number;
    totalActiveMin: number;
    activeDays: number;
    totalCards?: number;
    totalTokensSeen: number;
    totalLookupCount: number;
    totalLookupHits: number;
    totalYomitanLookupCount: number;
    newWordsToday: number;
    newWordsThisWeek: number;
  };
}

export interface MediaLibraryItem {
  videoId: number;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalTokensSeen: number;
  lastWatchedMs: number;
  hasCoverArt: number;
  youtubeVideoId?: string | null;
  videoUrl?: string | null;
  videoTitle?: string | null;
  videoThumbnailUrl?: string | null;
  channelId?: string | null;
  channelName?: string | null;
  channelUrl?: string | null;
  channelThumbnailUrl?: string | null;
  uploaderId?: string | null;
  uploaderUrl?: string | null;
  description?: string | null;
}

export interface MediaDetailData {
  detail: {
    videoId: number;
    canonicalTitle: string;
    animeId: number | null;
    totalSessions: number;
    totalActiveMs: number;
    totalCards: number;
    totalTokensSeen: number;
    totalLinesSeen: number;
    totalLookupCount: number;
    totalLookupHits: number;
    totalYomitanLookupCount: number;
    youtubeVideoId?: string | null;
    videoUrl?: string | null;
    videoTitle?: string | null;
    videoThumbnailUrl?: string | null;
    channelId?: string | null;
    channelName?: string | null;
    channelUrl?: string | null;
    channelThumbnailUrl?: string | null;
    uploaderId?: string | null;
    uploaderUrl?: string | null;
    description?: string | null;
  } | null;
  sessions: SessionSummary[];
  rollups: DailyRollup[];
}

export const EventType = {
  SUBTITLE_LINE: 1,
  MEDIA_BUFFER: 2,
  LOOKUP: 3,
  CARD_MINED: 4,
  SEEK_FORWARD: 5,
  SEEK_BACKWARD: 6,
  PAUSE_START: 7,
  PAUSE_END: 8,
  YOMITAN_LOOKUP: 9,
} as const;

export type EventType = (typeof EventType)[keyof typeof EventType];

export interface AnimeLibraryItem {
  mediaKind: MediaKind;
  animeId: number;
  canonicalTitle: string;
  anilistId: number | null;
  tmdbId: number | null;
  tmdbType: TmdbMediaType | null;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalTokensSeen: number;
  episodeCount: number;
  episodesTotal: number | null;
  lastWatchedMs: number;
}

export interface AnilistEntry {
  anilistId: number;
  titleRomaji: string | null;
  titleEnglish: string | null;
  season: number | null;
}

export interface AnimeDetailData {
  detail: {
    mediaKind: MediaKind;
    animeId: number;
    canonicalTitle: string;
    anilistId: number | null;
    tmdbId: number | null;
    tmdbType: TmdbMediaType | null;
    titleRomaji: string | null;
    titleEnglish: string | null;
    titleNative: string | null;
    description: string | null;
    totalSessions: number;
    totalActiveMs: number;
    totalCards: number;
    totalTokensSeen: number;
    totalLinesSeen: number;
    totalLookupCount: number;
    totalLookupHits: number;
    totalYomitanLookupCount: number;
    episodeCount: number;
    lastWatchedMs: number;
  };
  episodes: AnimeEpisode[];
  anilistEntries: AnilistEntry[];
}

export interface AnimeEpisode {
  videoId: number;
  episode: number | null;
  season: number | null;
  durationMs: number;
  endedMediaMs: number | null;
  watched: number;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalTokensSeen: number;
  totalYomitanLookupCount: number;
  lastWatchedMs: number;
}

export interface AnimeWord {
  wordId: number;
  headword: string;
  word: string;
  reading: string;
  partOfSpeech: string | null;
  frequency: number;
}

export interface StreakCalendarDay {
  epochDay: number;
  totalActiveMin: number;
}

export interface EpisodesPerDay {
  epochDay: number;
  episodeCount: number;
}

export interface NewAnimePerDay {
  epochDay: number;
  newAnimeCount: number;
}

export interface WatchTimePerAnime {
  epochDay: number;
  animeId: number;
  animeTitle: string;
  totalActiveMin: number;
}

export interface TrendChartPoint {
  label: string;
  value: number;
}

/** A rate over a bucket. `null` means too little data to rate, which charts draw as a gap. */
export interface TrendRatePoint {
  label: string;
  value: number | null;
}

/** One x-axis bucket, clipped to the requested date range. Epoch days are local calendar days. */
export interface TrendBucket {
  label: string;
  startDay: number;
  endDay: number;
}

export interface TrendPerAnimePoint {
  epochDay: number;
  animeTitle: string;
  value: number;
}

export const TREND_FREQUENCY_BANDS = [
  'top1k',
  'top5k',
  'top10k',
  'top20k',
  'beyond20k',
  'unranked',
] as const;
export type TrendFrequencyBand = (typeof TREND_FREQUENCY_BANDS)[number];

export interface TrendFrequencyBandPoint {
  label: string;
  counts: Record<TrendFrequencyBand, number>;
}

export interface TrendSummaryTotals {
  watchTimeMin: number;
  words: number;
  /** Null when the dashboard is filtered to one title: new words are not tracked per title. */
  newWords: number | null;
  cards: number;
  episodes: number;
  activeDays: number;
}

export interface LibrarySummaryRow {
  title: string;
  /** Null for videos that are not linked to a library title. */
  animeId: number | null;
  /** Most recently watched video, the click-through target when `animeId` is null. */
  videoId: number | null;
  watchTimeMin: number;
  videos: number;
  episodesFinished: number;
  sessions: number;
  cards: number;
  words: number;
  lookups: number;
  lookupsPerHundred: number | null;
  firstWatched: number;
  lastWatched: number;
}

export interface TrendsDashboardData {
  /** The days the dashboard covers. For `all` this is the span of recorded data. */
  window: { startDay: number; endDay: number };
  /** Axis shared by the `activity`, `progress`, and `ratios` series, index for index. */
  buckets: TrendBucket[];
  summary: {
    current: TrendSummaryTotals;
    /** The equally long period just before `window`; null for `all`. */
    previous: TrendSummaryTotals | null;
  };
  activity: {
    watchTime: TrendChartPoint[];
    cards: TrendChartPoint[];
    words: TrendChartPoint[];
    episodes: TrendChartPoint[];
  };
  /** Running totals counted from the start of `window`, each point the total at its period's end. */
  progress: {
    watchTime: TrendChartPoint[];
    words: TrendChartPoint[];
    newWords: TrendChartPoint[];
    cards: TrendChartPoint[];
    episodes: TrendChartPoint[];
    lookups: TrendChartPoint[];
  };
  ratios: {
    lookupsPerHundred: TrendRatePoint[];
    cardsPerHour: TrendRatePoint[];
    readingSpeed: TrendRatePoint[];
  };
  /** Vocabulary growth signals. Null when filtered to one title. */
  learning: {
    newWordsPerThousand: TrendRatePoint[];
    frequencyBands: TrendFrequencyBandPoint[];
  } | null;
  /** Active minutes as `[dayOfWeek][hour]`, with 0 = Sunday and local-time hours. */
  watchTimeByDayHour: number[][];
  /** Every title in `window`, unaffected by the title filter. */
  librarySummary: LibrarySummaryRow[];
  /** Sparse active minutes per title per day, unaffected by the title filter. */
  titleTimeline: TrendPerAnimePoint[];
}

export interface WordDetailData {
  detail: {
    wordId: number;
    headword: string;
    word: string;
    reading: string;
    partOfSpeech: string | null;
    pos1: string | null;
    pos2: string | null;
    pos3: string | null;
    /** Times seen across every form of this dictionary word. */
    frequency: number;
    frequencyRank: number | null;
    /** Reading of the dictionary form, or empty when only conjugated forms were seen. */
    dictionaryReading: string;
    formCount: number;
    forms: DictionaryWordForm[];
    /** Stored subtitle lines with this word, across every form. */
    lineCount: number;
    /** Those lines that came from a stream, which can't be mined. */
    streamLineCount: number;
    firstSeen: number;
    lastSeen: number;
  };
  animeAppearances: Array<{
    animeId: number;
    animeTitle: string;
    occurrenceCount: number;
  }>;
  similarWords: Array<{
    wordId: number;
    headword: string;
    word: string;
    reading: string;
    frequency: number;
  }>;
}

export interface EpisodeCardEvent {
  eventId: number;
  sessionId: number;
  tsMs: number;
  cardsDelta: number;
  noteIds: number[];
}

export interface EpisodeDetailData {
  sessions: SessionSummary[];
  words: AnimeWord[];
  cardEvents: EpisodeCardEvent[];
}

export interface KanjiDetailData {
  detail: {
    kanjiId: number;
    kanji: string;
    frequency: number;
    firstSeen: number;
    lastSeen: number;
  };
  animeAppearances: Array<{
    animeId: number;
    animeTitle: string;
    occurrenceCount: number;
  }>;
  words: Array<{
    wordId: number;
    headword: string;
    word: string;
    reading: string;
    frequency: number;
  }>;
}
