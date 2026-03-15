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
  wordsSeen: number;
  tokensSeen: number;
  cardsMined: number;
  lookupCount: number;
  lookupHits: number;
}

export interface DailyRollup {
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

export type MonthlyRollup = DailyRollup;

export interface SessionTimelinePoint {
  sampleMs: number;
  totalWatchedMs: number;
  activeWatchedMs: number;
  linesSeen: number;
  wordsSeen: number;
  tokensSeen: number;
  cardsMined: number;
}

export interface SessionEvent {
  eventType: EventType;
  tsMs: number;
  payload: string | null;
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
  firstSeen: number;
  lastSeen: number;
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
  sessionId: number;
  lineIndex: number;
  segmentStartMs: number | null;
  segmentEndMs: number | null;
  text: string;
  occurrenceCount: number;
}

export interface OverviewData {
  sessions: SessionSummary[];
  rollups: DailyRollup[];
  hints: {
    totalSessions: number;
    activeSessions: number;
    episodesToday: number;
    activeAnimeCount: number;
  };
}

export interface MediaLibraryItem {
  videoId: number;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
  totalWordsSeen: number;
  lastWatchedMs: number;
  hasCoverArt: number;
}

export interface MediaDetailData {
  detail: {
    videoId: number;
    canonicalTitle: string;
    totalSessions: number;
    totalActiveMs: number;
    totalCards: number;
    totalWordsSeen: number;
    totalLinesSeen: number;
    totalLookupCount: number;
    totalLookupHits: number;
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
} as const;

export type EventType = (typeof EventType)[keyof typeof EventType];

export interface AnimeLibraryItem {
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
    totalSessions: number;
    totalActiveMs: number;
    totalCards: number;
    totalWordsSeen: number;
    totalLinesSeen: number;
    totalLookupCount: number;
    totalLookupHits: number;
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
  watched: number;
  canonicalTitle: string;
  totalSessions: number;
  totalActiveMs: number;
  totalCards: number;
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
