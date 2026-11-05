import assert from 'node:assert/strict';
import test from 'node:test';
import {
  closeMediaDetail,
  createInitialStatsView,
  getSessionNavigationTarget,
  navigateToAnime,
  navigateToSessionsDays,
  openAnimeEpisodeDetail,
  openOverviewMediaDetail,
  openSessionsMediaDetail,
  openTrendsMediaDetail,
  switchTab,
  type StatsViewState,
} from './stats-navigation';

test('openAnimeEpisodeDetail opens dedicated media detail from anime context', () => {
  const state = createInitialStatsView();

  assert.deepEqual(openAnimeEpisodeDetail(state, 42, 7), {
    activeTab: 'anime',
    selectedAnimeId: 42,
    focusedSessionId: null,
    sessionsDayFilter: null,
    mediaDetail: {
      videoId: 7,
      initialSessionId: null,
      origin: {
        type: 'anime',
        animeId: 42,
      },
    },
  } satisfies StatsViewState);
});

test('closeMediaDetail returns to originating anime detail state', () => {
  const state = openAnimeEpisodeDetail(navigateToAnime(createInitialStatsView(), 42), 42, 7);

  assert.deepEqual(closeMediaDetail(state), {
    activeTab: 'anime',
    selectedAnimeId: 42,
    focusedSessionId: null,
    sessionsDayFilter: null,
    mediaDetail: null,
  } satisfies StatsViewState);
});

test('openOverviewMediaDetail opens dedicated media detail from overview context', () => {
  assert.deepEqual(openOverviewMediaDetail(createInitialStatsView(), 9), {
    activeTab: 'overview',
    selectedAnimeId: null,
    focusedSessionId: null,
    sessionsDayFilter: null,
    mediaDetail: {
      videoId: 9,
      initialSessionId: null,
      origin: {
        type: 'overview',
      },
    },
  } satisfies StatsViewState);
});

test('closeMediaDetail returns to overview when media detail originated there', () => {
  const state = openOverviewMediaDetail(createInitialStatsView(), 9);

  assert.deepEqual(closeMediaDetail(state), createInitialStatsView());
});

test('switchTab clears dedicated media detail state', () => {
  const state = openAnimeEpisodeDetail(navigateToAnime(createInitialStatsView(), 42), 42, 7);

  assert.deepEqual(switchTab(state, 'sessions'), {
    activeTab: 'sessions',
    selectedAnimeId: null,
    focusedSessionId: null,
    sessionsDayFilter: null,
    mediaDetail: null,
  } satisfies StatsViewState);
});

test('getSessionNavigationTarget prefers media detail when video id exists', () => {
  assert.deepEqual(getSessionNavigationTarget({ sessionId: 4, videoId: 12 }), {
    type: 'media-detail',
    videoId: 12,
    sessionId: 4,
  });
});

test('getSessionNavigationTarget falls back to session page when video id is missing', () => {
  assert.deepEqual(getSessionNavigationTarget({ sessionId: 4, videoId: null }), {
    type: 'session',
    sessionId: 4,
  });
});

test('openOverviewMediaDetail can carry a target session id for auto-expansion', () => {
  assert.deepEqual(openOverviewMediaDetail(createInitialStatsView(), 9, 33), {
    activeTab: 'overview',
    selectedAnimeId: null,
    focusedSessionId: null,
    sessionsDayFilter: null,
    mediaDetail: {
      videoId: 9,
      initialSessionId: 33,
      origin: {
        type: 'overview',
      },
    },
  } satisfies StatsViewState);
});

const OCT_5 = { startDay: 20_366, endDay: 20_366, label: 'Oct 5' };

test('navigateToSessionsDays opens Sessions narrowed to those days', () => {
  const state = navigateToSessionsDays(createInitialStatsView(), OCT_5);
  assert.equal(state.activeTab, 'sessions');
  assert.deepEqual(state.sessionsDayFilter, OCT_5);
});

test('the sessions day filter survives a media detail round trip but not a tab switch', () => {
  const filtered = navigateToSessionsDays(createInitialStatsView(), OCT_5);
  const back = closeMediaDetail(openSessionsMediaDetail(filtered, 3));
  assert.equal(back.activeTab, 'sessions');
  assert.deepEqual(back.sessionsDayFilter, OCT_5);
  assert.equal(switchTab(filtered, 'sessions').sessionsDayFilter, null);
});

test('media detail opened from Trends returns to Trends', () => {
  const state = closeMediaDetail(openTrendsMediaDetail(createInitialStatsView(), 3));
  assert.equal(state.activeTab, 'trends');
  assert.equal(state.mediaDetail, null);
});
