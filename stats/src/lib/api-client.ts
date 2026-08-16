import type {
  StatsAnkiNoteInfo,
  StatsExcludedWord,
  StatsCoverImagesData,
  StatsAnilistAssignment,
  StatsAnkiNotesInfoRequest,
  StatsCoverImagesRequest,
  StatsDeleteSessionsRequest,
  StatsDuplicateLineCleanupRequest,
  StatsDuplicateLineCleanupResult,
  StatsExcludedWordsRequest,
  StatsHttpClient,
  StatsJsonResponseMap,
  StatsMergeAnimeRequest,
  StatsMergeAnimeResponse,
  StatsMoveVideoRequest,
  StatsMoveVideoResponse,
  StatsTrendGroupBy,
  StatsTrendRange,
  StatsVideoWatchedRequest,
} from '../types/stats';
import type { StatsMineCardParams, StatsMineCardResponse } from './mining';
import { appendCoverRetryToken } from './cover-retry';
import { trackDelete } from './delete-progress';

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

async function fetchJson<Key extends keyof StatsJsonResponseMap>(
  _key: Key,
  path: string,
): Promise<StatsJsonResponseMap[Key]> {
  const res = await fetchResponse(path);
  return res.json() as Promise<StatsJsonResponseMap[Key]>;
}

function uniquePositiveIds(ids: number[]): number[] {
  const uniqueIds = new Set<number>();
  for (const id of ids) {
    if (Number.isFinite(id) && id > 0) {
      uniqueIds.add(Math.floor(id));
    }
  }
  return Array.from(uniqueIds).sort((a, b) => a - b);
}

export const apiClient = {
  getOverview: () => fetchJson('overview', '/api/stats/overview'),
  getDailyRollups: (limit = 60) =>
    fetchJson('dailyRollups', `/api/stats/daily-rollups?limit=${limit}`),
  getMonthlyRollups: (limit = 24) =>
    fetchJson('monthlyRollups', `/api/stats/monthly-rollups?limit=${limit}`),
  getSessions: (limit = 50) => fetchJson('sessions', `/api/stats/sessions?limit=${limit}`),
  getSessionTimeline: (id: number, limit?: number) =>
    fetchJson(
      'sessionTimeline',
      limit === undefined
        ? `/api/stats/sessions/${id}/timeline`
        : `/api/stats/sessions/${id}/timeline?limit=${limit}`,
    ),
  getSessionEvents: (id: number, limit = 500, eventTypes?: number[]) => {
    const params = new URLSearchParams({ limit: String(limit) });
    if (eventTypes && eventTypes.length > 0) {
      params.set('types', eventTypes.join(','));
    }
    return fetchJson('sessionEvents', `/api/stats/sessions/${id}/events?${params.toString()}`);
  },
  getSessionKnownWordsTimeline: (id: number) =>
    fetchJson('sessionKnownWordsTimeline', `/api/stats/sessions/${id}/known-words-timeline`),
  getVocabulary: (limit = 100) => fetchJson('vocabulary', `/api/stats/vocabulary?limit=${limit}`),
  getVocabularySummary: () => fetchJson('vocabularySummary', '/api/stats/vocabulary/summary'),
  getVocabularyCharts: () => fetchJson('vocabularyCharts', '/api/stats/vocabulary/charts'),
  getExcludedWords: () => fetchJson('excludedWords', '/api/stats/excluded-words'),
  setExcludedWords: async (words: StatsExcludedWord[]): Promise<void> => {
    await fetchResponse('/api/stats/excluded-words', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ words } satisfies StatsExcludedWordsRequest),
    });
  },
  cleanupDuplicateLines: async (
    options: StatsDuplicateLineCleanupRequest = {},
  ): Promise<StatsDuplicateLineCleanupResult> => {
    const res = await fetchResponse('/api/stats/maintenance/duplicate-lines', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        dryRun: options.dryRun === true,
        lookbackDays: options.lookbackDays ?? null,
      } satisfies StatsDuplicateLineCleanupRequest),
    });
    return res.json() as Promise<StatsDuplicateLineCleanupResult>;
  },
  getWordOccurrences: (headword: string, word: string, reading: string, limit = 50, offset = 0) =>
    fetchJson(
      'wordOccurrences',
      `/api/stats/vocabulary/occurrences?headword=${encodeURIComponent(headword)}&word=${encodeURIComponent(word)}&reading=${encodeURIComponent(reading)}&limit=${limit}&offset=${offset}`,
    ),
  searchSentences: (query: string, limit = 50, searchByHeadword = true) =>
    fetchJson(
      'sentenceSearch',
      `/api/stats/sentences/search?${new URLSearchParams({
        q: query,
        limit: String(limit),
        headword: String(searchByHeadword),
      }).toString()}`,
    ),
  getKanji: (limit = 100) => fetchJson('kanji', `/api/stats/kanji?limit=${limit}`),
  getKanjiOccurrences: (kanji: string, limit = 50, offset = 0) =>
    fetchJson(
      'kanjiOccurrences',
      `/api/stats/kanji/occurrences?kanji=${encodeURIComponent(kanji)}&limit=${limit}&offset=${offset}`,
    ),
  getMediaLibrary: () => fetchJson('mediaLibrary', '/api/stats/media'),
  getMediaDetail: (videoId: number) => fetchJson('mediaDetail', `/api/stats/media/${videoId}`),
  getAnimeLibrary: () => fetchJson('animeLibrary', '/api/stats/anime'),
  getAnimeMergeRecommendations: () =>
    fetchJson('animeMergeRecommendations', '/api/stats/anime/merge-recommendations'),
  getAnimeDetail: (animeId: number) => fetchJson('animeDetail', `/api/stats/anime/${animeId}`),
  getAnimeWords: (animeId: number, limit = 50) =>
    fetchJson('animeWords', `/api/stats/anime/${animeId}/words?limit=${limit}`),
  getAnimeRollups: (animeId: number, limit = 90) =>
    fetchJson('animeRollups', `/api/stats/anime/${animeId}/rollups?limit=${limit}`),
  getAnimeCoverUrl: (animeId: number, retryToken = 0) =>
    appendCoverRetryToken(`${BASE_URL}/api/stats/anime/${animeId}/cover`, retryToken),
  getCoverImages: async (params: StatsCoverImagesRequest): Promise<StatsCoverImagesData> => {
    const res = await fetchResponse('/api/stats/covers', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        animeIds: uniquePositiveIds(params.animeIds),
        videoIds: uniquePositiveIds(params.videoIds),
      } satisfies StatsCoverImagesRequest),
    });
    return res.json() as Promise<StatsJsonResponseMap['coverImages']>;
  },
  getStreakCalendar: (days = 90) =>
    fetchJson('streakCalendar', `/api/stats/streak-calendar?days=${days}`),
  getEpisodesPerDay: (limit = 90) =>
    fetchJson('episodesPerDay', `/api/stats/trends/episodes-per-day?limit=${limit}`),
  getNewAnimePerDay: (limit = 90) =>
    fetchJson('newAnimePerDay', `/api/stats/trends/new-anime-per-day?limit=${limit}`),
  getWatchTimePerAnime: (limit = 90) =>
    fetchJson('watchTimePerAnime', `/api/stats/trends/watch-time-per-anime?limit=${limit}`),
  getTrendsDashboard: (range: StatsTrendRange, groupBy: StatsTrendGroupBy, fillEmpty = true) =>
    fetchJson(
      'trendsDashboard',
      `/api/stats/trends/dashboard?range=${encodeURIComponent(range)}&groupBy=${encodeURIComponent(groupBy)}&fillEmpty=${fillEmpty ? 'true' : 'false'}`,
    ),
  getWordDetail: (wordId: number) =>
    fetchJson('wordDetail', `/api/stats/vocabulary/${wordId}/detail`),
  getKanjiDetail: (kanjiId: number) =>
    fetchJson('kanjiDetail', `/api/stats/kanji/${kanjiId}/detail`),
  getEpisodeDetail: (videoId: number) =>
    fetchJson('episodeDetail', `/api/stats/episode/${videoId}/detail`),
  setVideoWatched: async (videoId: number, watched: boolean): Promise<void> => {
    await fetchResponse(`/api/stats/media/${videoId}/watched`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ watched } satisfies StatsVideoWatchedRequest),
    });
  },
  deleteSession: async (sessionId: number): Promise<void> => {
    await trackDelete('Deleting session', () =>
      fetchResponse(`/api/stats/sessions/${sessionId}`, { method: 'DELETE' }),
    );
  },
  deleteSessions: async (sessionIds: number[]): Promise<void> => {
    const label = `Deleting ${sessionIds.length} session${sessionIds.length === 1 ? '' : 's'}`;
    await trackDelete(label, () =>
      fetchResponse('/api/stats/sessions', {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sessionIds } satisfies StatsDeleteSessionsRequest),
      }),
    );
  },
  deleteVideo: async (videoId: number): Promise<void> => {
    await trackDelete('Deleting episode', () =>
      fetchResponse(`/api/stats/media/${videoId}`, { method: 'DELETE' }),
    );
  },
  deleteAnime: async (animeId: number): Promise<void> => {
    await trackDelete('Deleting library entry', () =>
      fetchResponse(`/api/stats/anime/${animeId}`, { method: 'DELETE' }),
    );
  },
  mergeAnime: async (
    targetAnimeId: number,
    sourceAnimeIds: number[],
  ): Promise<StatsMergeAnimeResponse> => {
    const res = await fetchResponse(`/api/stats/anime/${targetAnimeId}/merge`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sourceAnimeIds } satisfies StatsMergeAnimeRequest),
    });
    return res.json() as Promise<StatsMergeAnimeResponse>;
  },
  moveVideoToAnime: async (videoId: number, animeId: number): Promise<StatsMoveVideoResponse> => {
    const res = await fetchResponse(`/api/stats/media/${videoId}/anime`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ animeId } satisfies StatsMoveVideoRequest),
    });
    return res.json() as Promise<StatsMoveVideoResponse>;
  },
  dismissAnimeMergeRecommendation: async (recommendationId: number): Promise<void> => {
    await fetchResponse(`/api/stats/anime/merge-recommendations/${recommendationId}`, {
      method: 'DELETE',
    });
  },
  getKnownWords: () => fetchJson('knownWords', '/api/stats/known-words'),
  getKnownWordsSummary: () => fetchJson('knownWordsSummary', '/api/stats/known-words-summary'),
  getAnimeKnownWordsSummary: (animeId: number) =>
    fetchJson('animeKnownWordsSummary', `/api/stats/anime/${animeId}/known-words-summary`),
  getMediaKnownWordsSummary: (videoId: number) =>
    fetchJson('mediaKnownWordsSummary', `/api/stats/media/${videoId}/known-words-summary`),
  searchAnilist: (query: string) =>
    fetchJson('anilistSearch', `/api/stats/anilist/search?q=${encodeURIComponent(query)}`),
  reassignAnimeAnilist: async (animeId: number, info: StatsAnilistAssignment): Promise<void> => {
    await fetchResponse(`/api/stats/anime/${animeId}/anilist`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(info),
    });
  },
  mineCard: async (params: StatsMineCardParams): Promise<StatsMineCardResponse> => {
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
  ankiNotesInfo: async (noteIds: number[]): Promise<StatsAnkiNoteInfo[]> => {
    const res = await fetch(`${BASE_URL}/api/stats/anki/notesInfo`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ noteIds } satisfies StatsAnkiNotesInfoRequest),
    });
    if (!res.ok) throw new Error(`Stats API error: ${res.status}`);
    return res.json();
  },
} satisfies StatsHttpClient;
