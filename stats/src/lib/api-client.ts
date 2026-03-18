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
  TrendsDashboardData,
  WordDetailData,
  KanjiDetailData,
  EpisodeDetailData,
} from '../types/stats';

type StatsLocationLike = Pick<Location, 'protocol' | 'origin' | 'search'>;

export function resolveStatsBaseUrl(location?: StatsLocationLike): string {
  const resolvedLocation =
    location ??
    (typeof window === 'undefined'
      ? { protocol: 'file:', origin: 'null', search: '' }
      : window.location);

  const queryApiBase = new URLSearchParams(resolvedLocation.search).get('apiBase')?.trim();
  if (queryApiBase) {
    return queryApiBase;
  }

  return resolvedLocation.protocol === 'file:' ? 'http://127.0.0.1:6969' : resolvedLocation.origin;
}

export const BASE_URL = resolveStatsBaseUrl();

async function fetchResponse(path: string, init?: RequestInit): Promise<Response> {
  const res = await fetch(`${BASE_URL}${path}`, init);
  if (!res.ok) {
    let body = '';
    try {
      body = (await res.text()).trim();
    } catch {
      body = '';
    }
    throw new Error(
      body ? `Stats API error: ${res.status} ${body}` : `Stats API error: ${res.status}`,
    );
  }
  return res;
}

async function fetchJson<T>(path: string): Promise<T> {
  const res = await fetchResponse(path);
  return res.json() as Promise<T>;
}

export const apiClient = {
  getOverview: () => fetchJson<OverviewData>('/api/stats/overview'),
  getDailyRollups: (limit = 60) =>
    fetchJson<DailyRollup[]>(`/api/stats/daily-rollups?limit=${limit}`),
  getMonthlyRollups: (limit = 24) =>
    fetchJson<MonthlyRollup[]>(`/api/stats/monthly-rollups?limit=${limit}`),
  getSessions: (limit = 50) => fetchJson<SessionSummary[]>(`/api/stats/sessions?limit=${limit}`),
  getSessionTimeline: (id: number, limit?: number) =>
    fetchJson<SessionTimelinePoint[]>(
      limit === undefined
        ? `/api/stats/sessions/${id}/timeline`
        : `/api/stats/sessions/${id}/timeline?limit=${limit}`,
    ),
  getSessionEvents: (id: number, limit = 500) =>
    fetchJson<SessionEvent[]>(`/api/stats/sessions/${id}/events?limit=${limit}`),
  getSessionKnownWordsTimeline: (id: number) =>
    fetchJson<Array<{ linesSeen: number; knownWordsSeen: number }>>(
      `/api/stats/sessions/${id}/known-words-timeline`,
    ),
  getVocabulary: (limit = 100) =>
    fetchJson<VocabularyEntry[]>(`/api/stats/vocabulary?limit=${limit}`),
  getWordOccurrences: (headword: string, word: string, reading: string, limit = 50, offset = 0) =>
    fetchJson<VocabularyOccurrenceEntry[]>(
      `/api/stats/vocabulary/occurrences?headword=${encodeURIComponent(headword)}&word=${encodeURIComponent(word)}&reading=${encodeURIComponent(reading)}&limit=${limit}&offset=${offset}`,
    ),
  getKanji: (limit = 100) => fetchJson<KanjiEntry[]>(`/api/stats/kanji?limit=${limit}`),
  getKanjiOccurrences: (kanji: string, limit = 50, offset = 0) =>
    fetchJson<VocabularyOccurrenceEntry[]>(
      `/api/stats/kanji/occurrences?kanji=${encodeURIComponent(kanji)}&limit=${limit}&offset=${offset}`,
    ),
  getMediaLibrary: () => fetchJson<MediaLibraryItem[]>('/api/stats/media'),
  getMediaDetail: (videoId: number) => fetchJson<MediaDetailData>(`/api/stats/media/${videoId}`),
  getAnimeLibrary: () => fetchJson<AnimeLibraryItem[]>('/api/stats/anime'),
  getAnimeDetail: (animeId: number) => fetchJson<AnimeDetailData>(`/api/stats/anime/${animeId}`),
  getAnimeWords: (animeId: number, limit = 50) =>
    fetchJson<AnimeWord[]>(`/api/stats/anime/${animeId}/words?limit=${limit}`),
  getAnimeRollups: (animeId: number, limit = 90) =>
    fetchJson<DailyRollup[]>(`/api/stats/anime/${animeId}/rollups?limit=${limit}`),
  getAnimeCoverUrl: (animeId: number) => `${BASE_URL}/api/stats/anime/${animeId}/cover`,
  getStreakCalendar: (days = 90) =>
    fetchJson<StreakCalendarDay[]>(`/api/stats/streak-calendar?days=${days}`),
  getEpisodesPerDay: (limit = 90) =>
    fetchJson<EpisodesPerDay[]>(`/api/stats/trends/episodes-per-day?limit=${limit}`),
  getNewAnimePerDay: (limit = 90) =>
    fetchJson<NewAnimePerDay[]>(`/api/stats/trends/new-anime-per-day?limit=${limit}`),
  getWatchTimePerAnime: (limit = 90) =>
    fetchJson<WatchTimePerAnime[]>(`/api/stats/trends/watch-time-per-anime?limit=${limit}`),
  getTrendsDashboard: (range: '7d' | '30d' | '90d' | 'all', groupBy: 'day' | 'month') =>
    fetchJson<TrendsDashboardData>(
      `/api/stats/trends/dashboard?range=${encodeURIComponent(range)}&groupBy=${encodeURIComponent(groupBy)}`,
    ),
  getWordDetail: (wordId: number) =>
    fetchJson<WordDetailData>(`/api/stats/vocabulary/${wordId}/detail`),
  getKanjiDetail: (kanjiId: number) =>
    fetchJson<KanjiDetailData>(`/api/stats/kanji/${kanjiId}/detail`),
  getEpisodeDetail: (videoId: number) =>
    fetchJson<EpisodeDetailData>(`/api/stats/episode/${videoId}/detail`),
  setVideoWatched: async (videoId: number, watched: boolean): Promise<void> => {
    await fetchResponse(`/api/stats/media/${videoId}/watched`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ watched }),
    });
  },
  deleteSession: async (sessionId: number): Promise<void> => {
    await fetchResponse(`/api/stats/sessions/${sessionId}`, { method: 'DELETE' });
  },
  deleteSessions: async (sessionIds: number[]): Promise<void> => {
    await fetchResponse('/api/stats/sessions', {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionIds }),
    });
  },
  deleteVideo: async (videoId: number): Promise<void> => {
    await fetchResponse(`/api/stats/media/${videoId}`, { method: 'DELETE' });
  },
  getKnownWords: () => fetchJson<string[]>('/api/stats/known-words'),
  getKnownWordsSummary: () =>
    fetchJson<{ totalUniqueWords: number; knownWordCount: number }>('/api/stats/known-words-summary'),
  getAnimeKnownWordsSummary: (animeId: number) =>
    fetchJson<{ totalUniqueWords: number; knownWordCount: number }>(
      `/api/stats/anime/${animeId}/known-words-summary`,
    ),
  getMediaKnownWordsSummary: (videoId: number) =>
    fetchJson<{ totalUniqueWords: number; knownWordCount: number }>(
      `/api/stats/media/${videoId}/known-words-summary`,
    ),
  searchAnilist: (query: string) =>
    fetchJson<
      Array<{
        id: number;
        episodes: number | null;
        season: string | null;
        seasonYear: number | null;
        coverImage: { large: string | null; medium: string | null } | null;
        title: { romaji: string | null; english: string | null; native: string | null } | null;
      }>
    >(`/api/stats/anilist/search?q=${encodeURIComponent(query)}`),
  reassignAnimeAnilist: async (
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
  ): Promise<void> => {
    await fetchResponse(`/api/stats/anime/${animeId}/anilist`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(info),
    });
  },
  mineCard: async (params: {
    sourcePath: string;
    startMs: number;
    endMs: number;
    sentence: string;
    word: string;
    secondaryText?: string | null;
    videoTitle: string;
    mode: 'word' | 'sentence' | 'audio';
  }): Promise<{ noteId?: number; error?: string; errors?: string[] }> => {
    const res = await fetch(`${BASE_URL}/api/stats/mine-card?mode=${params.mode}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(params),
    });
    return res.json();
  },
  ankiBrowse: async (noteId: number): Promise<void> => {
    await fetchResponse(`/api/stats/anki/browse?noteId=${noteId}`, { method: 'POST' });
  },
  ankiNotesInfo: async (
    noteIds: number[],
  ): Promise<Array<{ noteId: number; fields: Record<string, { value: string }> }>> => {
    const res = await fetch(`${BASE_URL}/api/stats/anki/notesInfo`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ noteIds }),
    });
    if (!res.ok) throw new Error(`Stats API error: ${res.status}`);
    return res.json();
  },
};
