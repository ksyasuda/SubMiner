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
  eventType: EventType;
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

export interface StatsExcludedWord {
  headword: string;
  word: string;
  reading: string;
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
  animeId: number;
  canonicalTitle: string;
  anilistId: number | null;
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

export interface TrendPerAnimePoint {
  epochDay: number;
  animeTitle: string;
  value: number;
}

export interface LibrarySummaryRow {
  title: string;
  watchTimeMin: number;
  videos: number;
  sessions: number;
  cards: number;
  words: number;
  lookups: number;
  lookupsPerHundred: number | null;
  firstWatched: number;
  lastWatched: number;
}

export interface TrendsDashboardData {
  activity: {
    watchTime: TrendChartPoint[];
    cards: TrendChartPoint[];
    words: TrendChartPoint[];
    sessions: TrendChartPoint[];
  };
  progress: {
    watchTime: TrendChartPoint[];
    sessions: TrendChartPoint[];
    words: TrendChartPoint[];
    newWords: TrendChartPoint[];
    cards: TrendChartPoint[];
    episodes: TrendChartPoint[];
    lookups: TrendChartPoint[];
  };
  ratios: {
    lookupsPerHundred: TrendChartPoint[];
  };
  librarySummary: LibrarySummaryRow[];
  animeCumulative: {
    watchTime: TrendPerAnimePoint[];
    episodes: TrendPerAnimePoint[];
    cards: TrendPerAnimePoint[];
    words: TrendPerAnimePoint[];
  };
  patterns: {
    watchTimeByDayOfWeek: TrendChartPoint[];
    watchTimeByHour: TrendChartPoint[];
  };
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
    frequency: number;
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
