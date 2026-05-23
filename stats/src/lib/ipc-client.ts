import type {
  OverviewData,
  DailyRollup,
  MonthlyRollup,
  SessionSummary,
  SessionTimelinePoint,
  SessionEvent,
  VocabularyEntry,
  KanjiEntry,
  VocabularyOccurrenceEntry,
  MediaLibraryItem,
  MediaDetailData,
  AnimeLibraryItem,
  AnimeDetailData,
  AnimeWord,
  StreakCalendarDay,
  EpisodesPerDay,
  NewAnimePerDay,
  WatchTimePerAnime,
  WordDetailData,
  KanjiDetailData,
  EpisodeDetailData,
  StatsAnkiNoteInfo,
} from '../types/stats';

interface StatsElectronAPI {
  stats: {
    getOverview: () => Promise<OverviewData>;
    getDailyRollups: (limit?: number) => Promise<DailyRollup[]>;
    getMonthlyRollups: (limit?: number) => Promise<MonthlyRollup[]>;
    getSessions: (limit?: number) => Promise<SessionSummary[]>;
    getSessionTimeline: (id: number, limit?: number) => Promise<SessionTimelinePoint[]>;
    getSessionEvents: (id: number, limit?: number) => Promise<SessionEvent[]>;
    getVocabulary: (limit?: number) => Promise<VocabularyEntry[]>;
    getWordOccurrences: (
      headword: string,
      word: string,
      reading: string,
      limit?: number,
      offset?: number,
    ) => Promise<VocabularyOccurrenceEntry[]>;
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
    getAnimeCoverUrl: (animeId: number) => string;
    getStreakCalendar: (days?: number) => Promise<StreakCalendarDay[]>;
    getEpisodesPerDay: (limit?: number) => Promise<EpisodesPerDay[]>;
    getNewAnimePerDay: (limit?: number) => Promise<NewAnimePerDay[]>;
    getWatchTimePerAnime: (limit?: number) => Promise<WatchTimePerAnime[]>;
    getWordDetail: (wordId: number) => Promise<WordDetailData>;
    getKanjiDetail: (kanjiId: number) => Promise<KanjiDetailData>;
    getEpisodeDetail: (videoId: number) => Promise<EpisodeDetailData>;
    ankiBrowse: (noteId: number) => Promise<void>;
    ankiNotesInfo: (noteIds: number[]) => Promise<StatsAnkiNoteInfo[]>;
    hideOverlay: () => void;
    confirmNativeDialog?: (message: string) => boolean;
    beginNativeDialog?: () => void;
    endNativeDialog?: () => void;
  };
}

declare global {
  interface Window {
    electronAPI?: StatsElectronAPI;
  }
}

function getIpc(): StatsElectronAPI['stats'] {
  const api = window.electronAPI?.stats;
  if (!api) throw new Error('Electron IPC not available');
  return api;
}

export const ipcClient = {
  getOverview: () => getIpc().getOverview(),
  getDailyRollups: (limit = 60) => getIpc().getDailyRollups(limit),
  getMonthlyRollups: (limit = 24) => getIpc().getMonthlyRollups(limit),
  getSessions: (limit = 50) => getIpc().getSessions(limit),
  getSessionTimeline: (id: number, limit?: number) => getIpc().getSessionTimeline(id, limit),
  getSessionEvents: (id: number, limit = 500) => getIpc().getSessionEvents(id, limit),
  getVocabulary: (limit = 100) => getIpc().getVocabulary(limit),
  getWordOccurrences: (headword: string, word: string, reading: string, limit = 50, offset = 0) =>
    getIpc().getWordOccurrences(headword, word, reading, limit, offset),
  getKanji: (limit = 100) => getIpc().getKanji(limit),
  getKanjiOccurrences: (kanji: string, limit = 50, offset = 0) =>
    getIpc().getKanjiOccurrences(kanji, limit, offset),
  getMediaLibrary: () => getIpc().getMediaLibrary(),
  getMediaDetail: (videoId: number) => getIpc().getMediaDetail(videoId),
  getAnimeLibrary: () => getIpc().getAnimeLibrary(),
  getAnimeDetail: (animeId: number) => getIpc().getAnimeDetail(animeId),
  getAnimeWords: (animeId: number, limit = 50) => getIpc().getAnimeWords(animeId, limit),
  getAnimeRollups: (animeId: number, limit = 90) => getIpc().getAnimeRollups(animeId, limit),
  getAnimeCoverUrl: (animeId: number) => getIpc().getAnimeCoverUrl(animeId),
  getStreakCalendar: (days = 90) => getIpc().getStreakCalendar(days),
  getEpisodesPerDay: (limit = 90) => getIpc().getEpisodesPerDay(limit),
  getNewAnimePerDay: (limit = 90) => getIpc().getNewAnimePerDay(limit),
  getWatchTimePerAnime: (limit = 90) => getIpc().getWatchTimePerAnime(limit),
  getWordDetail: (wordId: number) => getIpc().getWordDetail(wordId),
  getKanjiDetail: (kanjiId: number) => getIpc().getKanjiDetail(kanjiId),
  getEpisodeDetail: (videoId: number) => getIpc().getEpisodeDetail(videoId),
  ankiBrowse: (noteId: number) => getIpc().ankiBrowse(noteId),
  ankiNotesInfo: (noteIds: number[]) => getIpc().ankiNotesInfo(noteIds),
};
