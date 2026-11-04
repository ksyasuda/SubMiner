import type { SessionSummary, StatsDayRange } from '../types/stats';
import type { TabId } from '../components/layout/TabBar';

export type MediaDetailOrigin =
  | { type: 'anime'; animeId: number }
  | { type: 'overview' }
  | { type: 'sessions' }
  | { type: 'trends' };

export interface MediaDetailState {
  videoId: number;
  initialSessionId: number | null;
  origin: MediaDetailOrigin;
}

/** Narrows the Sessions tab to a span of days, set when arriving from a Trends chart. */
export interface SessionsDayFilter extends StatsDayRange {
  label: string;
}

export interface StatsViewState {
  activeTab: TabId;
  selectedAnimeId: number | null;
  focusedSessionId: number | null;
  sessionsDayFilter: SessionsDayFilter | null;
  mediaDetail: MediaDetailState | null;
}

export function createInitialStatsView(): StatsViewState {
  return {
    activeTab: 'overview',
    selectedAnimeId: null,
    focusedSessionId: null,
    sessionsDayFilter: null,
    mediaDetail: null,
  };
}

export function switchTab(state: StatsViewState, tabId: TabId): StatsViewState {
  return {
    activeTab: tabId,
    selectedAnimeId: null,
    focusedSessionId: tabId === 'sessions' ? state.focusedSessionId : null,
    sessionsDayFilter: null,
    mediaDetail: null,
  };
}

export function navigateToAnime(state: StatsViewState, animeId: number): StatsViewState {
  return {
    ...state,
    activeTab: 'anime',
    selectedAnimeId: animeId,
    mediaDetail: null,
  };
}

export function navigateToSession(state: StatsViewState, sessionId: number): StatsViewState {
  return {
    ...state,
    activeTab: 'sessions',
    focusedSessionId: sessionId,
    sessionsDayFilter: null,
    mediaDetail: null,
  };
}

export function navigateToSessionsDays(
  state: StatsViewState,
  sessionsDayFilter: SessionsDayFilter,
): StatsViewState {
  return {
    ...state,
    activeTab: 'sessions',
    focusedSessionId: null,
    sessionsDayFilter,
    mediaDetail: null,
  };
}

export function openAnimeEpisodeDetail(
  state: StatsViewState,
  animeId: number,
  videoId: number,
  sessionId: number | null = null,
): StatsViewState {
  return {
    activeTab: 'anime',
    selectedAnimeId: animeId,
    focusedSessionId: null,
    sessionsDayFilter: null,
    mediaDetail: {
      videoId,
      initialSessionId: sessionId,
      origin: {
        type: 'anime',
        animeId,
      },
    },
  };
}

export function openOverviewMediaDetail(
  state: StatsViewState,
  videoId: number,
  sessionId: number | null = null,
): StatsViewState {
  return {
    activeTab: 'overview',
    selectedAnimeId: null,
    focusedSessionId: null,
    sessionsDayFilter: null,
    mediaDetail: {
      videoId,
      initialSessionId: sessionId,
      origin: {
        type: 'overview',
      },
    },
  };
}

export function openSessionsMediaDetail(state: StatsViewState, videoId: number): StatsViewState {
  return {
    activeTab: 'sessions',
    selectedAnimeId: null,
    focusedSessionId: null,
    sessionsDayFilter: state.sessionsDayFilter,
    mediaDetail: {
      videoId,
      initialSessionId: null,
      origin: {
        type: 'sessions',
      },
    },
  };
}

export function openTrendsMediaDetail(_state: StatsViewState, videoId: number): StatsViewState {
  return {
    activeTab: 'trends',
    selectedAnimeId: null,
    focusedSessionId: null,
    sessionsDayFilter: null,
    mediaDetail: {
      videoId,
      initialSessionId: null,
      origin: {
        type: 'trends',
      },
    },
  };
}

export function closeMediaDetail(state: StatsViewState): StatsViewState {
  if (!state.mediaDetail) {
    return state;
  }

  const { origin } = state.mediaDetail;
  if (origin.type === 'anime') {
    return {
      activeTab: 'anime',
      selectedAnimeId: origin.animeId,
      focusedSessionId: null,
      sessionsDayFilter: null,
      mediaDetail: null,
    };
  }

  return {
    activeTab: origin.type,
    selectedAnimeId: null,
    focusedSessionId: null,
    sessionsDayFilter: origin.type === 'sessions' ? state.sessionsDayFilter : null,
    mediaDetail: null,
  };
}

export function getSessionNavigationTarget(session: Pick<SessionSummary, 'sessionId' | 'videoId'>):
  | {
      type: 'media-detail';
      videoId: number;
      sessionId: number;
    }
  | {
      type: 'session';
      sessionId: number;
    } {
  if (session.videoId != null) {
    return {
      type: 'media-detail',
      videoId: session.videoId,
      sessionId: session.sessionId,
    };
  }

  return {
    type: 'session',
    sessionId: session.sessionId,
  };
}
