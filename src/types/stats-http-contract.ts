import type {
  AnimeDetailData,
  AnimeLibraryItem,
  AnimeWord,
  DailyRollup,
  EpisodeDetailData,
  EpisodesPerDay,
  KanjiDetailData,
  KanjiEntry,
  MediaDetailData,
  MediaLibraryItem,
  MonthlyRollup,
  NewAnimePerDay,
  OverviewData,
  SentenceSearchResult,
  SessionEvent,
  SessionSummary,
  SessionTimelinePoint,
  StatsAnkiNoteInfo,
  StatsCoverImagesData,
  StatsDuplicateLineCleanupResult,
  StatsExcludedWord,
  StreakCalendarDay,
  TrendsDashboardData,
  VocabularyEntry,
  VocabularyOccurrenceEntry,
  WatchTimePerAnime,
  WordDetailData,
} from './stats-wire';

export type StatsTrendRange = '7d' | '30d' | '90d' | '365d' | 'all';
export type StatsTrendGroupBy = 'day' | 'month';
export type StatsMineMode = 'word' | 'sentence' | 'audio';

/** Body of `POST /api/stats/maintenance/duplicate-lines`. */
export interface StatsDuplicateLineCleanupRequest {
  dryRun?: boolean;
  /** Null means every recorded line, whatever its age. */
  lookbackDays?: number | null;
}

export interface StatsSessionKnownWordsTimelinePoint {
  linesSeen: number;
  knownWordsSeen: number;
  totalWordsSeen: number;
}

export interface StatsKnownWordsSummary {
  totalUniqueWords: number;
  knownWordCount: number;
}

export interface StatsAnilistSearchResult {
  id: number;
  episodes: number | null;
  season: string | null;
  seasonYear: number | null;
  description: string | null;
  coverImage: { large: string | null; medium: string | null } | null;
  title: { romaji: string | null; english: string | null; native: string | null } | null;
}

export interface StatsAnilistAssignment {
  anilistId: number;
  titleRomaji?: string | null;
  titleEnglish?: string | null;
  titleNative?: string | null;
  episodesTotal?: number | null;
  description?: string | null;
  coverUrl?: string | null;
}

export interface StatsMineCardParams {
  sourcePath: string;
  startMs: number;
  endMs: number;
  sentence: string;
  word: string;
  secondaryText?: string | null;
  videoTitle: string;
  mode: StatsMineMode;
}

export interface StatsMineCardResponse {
  noteId?: number;
  error?: string;
  errors?: string[];
}

export interface StatsCoverImagesRequest {
  animeIds: number[];
  videoIds: number[];
}

export interface StatsExcludedWordsRequest {
  words: StatsExcludedWord[];
}

export interface StatsVideoWatchedRequest {
  watched: boolean;
}

export interface StatsDeleteSessionsRequest {
  sessionIds: number[];
}

export interface StatsAnkiNotesInfoRequest {
  noteIds: number[];
}

export interface StatsOkResponse {
  ok: true;
}

export interface StatsErrorResponse {
  error: string;
}

export interface StatsAnkiBrowseResponse {
  result?: number[] | null;
  error?: string | null;
}

export interface StatsJsonResponseMap {
  error: StatsErrorResponse | null;
  overview: OverviewData;
  dailyRollups: DailyRollup[];
  monthlyRollups: MonthlyRollup[];
  sessions: SessionSummary[];
  sessionTimeline: SessionTimelinePoint[];
  sessionEvents: SessionEvent[];
  sessionKnownWordsTimeline: StatsSessionKnownWordsTimelinePoint[];
  vocabulary: VocabularyEntry[];
  excludedWords: StatsExcludedWord[];
  setExcludedWords: StatsOkResponse;
  duplicateLineCleanup: StatsDuplicateLineCleanupResult;
  wordOccurrences: VocabularyOccurrenceEntry[];
  sentenceSearch: SentenceSearchResult[];
  kanji: KanjiEntry[];
  kanjiOccurrences: VocabularyOccurrenceEntry[];
  wordDetail: WordDetailData;
  kanjiDetail: KanjiDetailData;
  mediaLibrary: MediaLibraryItem[];
  mediaDetail: MediaDetailData;
  animeLibrary: AnimeLibraryItem[];
  animeDetail: AnimeDetailData;
  animeWords: AnimeWord[];
  animeRollups: DailyRollup[];
  setVideoWatched: StatsOkResponse;
  deleteSessions: StatsOkResponse;
  deleteSession: StatsOkResponse;
  deleteVideo: StatsOkResponse;
  deleteAnime: StatsOkResponse;
  anilistSearch: StatsAnilistSearchResult[];
  knownWords: string[];
  knownWordsSummary: StatsKnownWordsSummary;
  animeKnownWordsSummary: StatsKnownWordsSummary;
  mediaKnownWordsSummary: StatsKnownWordsSummary;
  reassignAnimeAnilist: StatsOkResponse;
  coverImages: StatsCoverImagesData;
  episodeDetail: EpisodeDetailData;
  ankiBrowse: StatsAnkiBrowseResponse;
  ankiNotesInfo: StatsAnkiNoteInfo[];
  mineCard: StatsMineCardResponse;
  streakCalendar: StreakCalendarDay[];
  episodesPerDay: EpisodesPerDay[];
  newAnimePerDay: NewAnimePerDay[];
  watchTimePerAnime: WatchTimePerAnime[];
  trendsDashboard: TrendsDashboardData;
}

export function statsJson<Key extends keyof StatsJsonResponseMap>(
  _key: Key,
  payload: StatsJsonResponseMap[Key],
): StatsJsonResponseMap[Key] {
  return payload;
}

export interface StatsHttpClient {
  getOverview: () => Promise<OverviewData>;
  getDailyRollups: (limit?: number) => Promise<DailyRollup[]>;
  getMonthlyRollups: (limit?: number) => Promise<MonthlyRollup[]>;
  getSessions: (limit?: number) => Promise<SessionSummary[]>;
  getSessionTimeline: (id: number, limit?: number) => Promise<SessionTimelinePoint[]>;
  getSessionEvents: (id: number, limit?: number, eventTypes?: number[]) => Promise<SessionEvent[]>;
  getSessionKnownWordsTimeline: (id: number) => Promise<StatsSessionKnownWordsTimelinePoint[]>;
  getVocabulary: (limit?: number) => Promise<VocabularyEntry[]>;
  getExcludedWords: () => Promise<StatsExcludedWord[]>;
  setExcludedWords: (words: StatsExcludedWord[]) => Promise<void>;
  cleanupDuplicateLines: (
    options?: StatsDuplicateLineCleanupRequest,
  ) => Promise<StatsDuplicateLineCleanupResult>;
  getWordOccurrences: (
    headword: string,
    word: string,
    reading: string,
    limit?: number,
    offset?: number,
  ) => Promise<VocabularyOccurrenceEntry[]>;
  searchSentences: (
    query: string,
    limit?: number,
    searchByHeadword?: boolean,
  ) => Promise<SentenceSearchResult[]>;
  getKanji: (limit?: number) => Promise<KanjiEntry[]>;
  getKanjiOccurrences: (
    kanji: string,
    limit?: number,
    offset?: number,
  ) => Promise<VocabularyOccurrenceEntry[]>;
  getMediaLibrary: () => Promise<MediaLibraryItem[]>;
  getMediaDetail: (videoId: number) => Promise<MediaDetailData>;
  getAnimeLibrary: () => Promise<AnimeLibraryItem[]>;
  getAnimeDetail: (animeId: number) => Promise<AnimeDetailData>;
  getAnimeWords: (animeId: number, limit?: number) => Promise<AnimeWord[]>;
  getAnimeRollups: (animeId: number, limit?: number) => Promise<DailyRollup[]>;
  getAnimeCoverUrl: (animeId: number, retryToken?: number) => string;
  getCoverImages: (params: StatsCoverImagesRequest) => Promise<StatsCoverImagesData>;
  getStreakCalendar: (days?: number) => Promise<StreakCalendarDay[]>;
  getEpisodesPerDay: (limit?: number) => Promise<EpisodesPerDay[]>;
  getNewAnimePerDay: (limit?: number) => Promise<NewAnimePerDay[]>;
  getWatchTimePerAnime: (limit?: number) => Promise<WatchTimePerAnime[]>;
  getTrendsDashboard: (
    range: StatsTrendRange,
    groupBy: StatsTrendGroupBy,
    fillEmpty?: boolean,
  ) => Promise<TrendsDashboardData>;
  getWordDetail: (wordId: number) => Promise<WordDetailData>;
  getKanjiDetail: (kanjiId: number) => Promise<KanjiDetailData>;
  getEpisodeDetail: (videoId: number) => Promise<EpisodeDetailData>;
  setVideoWatched: (videoId: number, watched: boolean) => Promise<void>;
  deleteSession: (sessionId: number) => Promise<void>;
  deleteSessions: (sessionIds: number[]) => Promise<void>;
  deleteVideo: (videoId: number) => Promise<void>;
  deleteAnime: (animeId: number) => Promise<void>;
  getKnownWords: () => Promise<string[]>;
  getKnownWordsSummary: () => Promise<StatsKnownWordsSummary>;
  getAnimeKnownWordsSummary: (animeId: number) => Promise<StatsKnownWordsSummary>;
  getMediaKnownWordsSummary: (videoId: number) => Promise<StatsKnownWordsSummary>;
  searchAnilist: (query: string) => Promise<StatsAnilistSearchResult[]>;
  reassignAnimeAnilist: (animeId: number, info: StatsAnilistAssignment) => Promise<void>;
  mineCard: (params: StatsMineCardParams) => Promise<StatsMineCardResponse>;
  ankiBrowse: (noteId: number) => Promise<void>;
  ankiNotesInfo: (noteIds: number[]) => Promise<StatsAnkiNoteInfo[]>;
}
