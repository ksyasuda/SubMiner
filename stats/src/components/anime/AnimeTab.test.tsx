import assert from 'node:assert/strict';
import test from 'node:test';
import { Window } from 'happy-dom';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { apiClient } from '../../lib/api-client';
import type { AnimeDetailData, AnimeLibraryItem } from '../../types/stats';
import { AnimeTab } from './AnimeTab';

interface TestWindow extends Window {
  IS_REACT_ACT_ENVIRONMENT?: boolean;
}

function installDom(): () => void {
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;
  const previousHTMLElement = globalThis.HTMLElement;
  const previousIsReactActEnvironment = (
    globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
  ).IS_REACT_ACT_ENVIRONMENT;
  const window = new Window() as TestWindow;

  Object.defineProperty(globalThis, 'window', { value: window, configurable: true });
  Object.defineProperty(globalThis, 'document', { value: window.document, configurable: true });
  Object.defineProperty(globalThis, 'HTMLElement', {
    value: window.HTMLElement,
    configurable: true,
  });
  (
    globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
  ).IS_REACT_ACT_ENVIRONMENT = true;

  return () => {
    Object.defineProperty(globalThis, 'window', { value: previousWindow, configurable: true });
    Object.defineProperty(globalThis, 'document', { value: previousDocument, configurable: true });
    Object.defineProperty(globalThis, 'HTMLElement', {
      value: previousHTMLElement,
      configurable: true,
    });
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = previousIsReactActEnvironment;
  };
}

function libraryItem(anilistId: number | null): AnimeLibraryItem {
  return {
    mediaKind: 'anime',
    animeId: 7,
    canonicalTitle: 'Test Anime Season 2',
    anilistId,
    tmdbId: null,
    tmdbType: null,
    totalSessions: 1,
    totalActiveMs: 1000,
    totalCards: 0,
    totalTokensSeen: 0,
    episodeCount: 1,
    episodesTotal: 13,
    lastWatchedMs: 1,
  };
}

function detailData(anilistId: number | null): AnimeDetailData {
  return {
    detail: {
      mediaKind: 'anime',
      animeId: 7,
      canonicalTitle: 'Test Anime Season 2',
      anilistId,
      tmdbId: null,
      tmdbType: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      description: null,
      totalSessions: 1,
      totalActiveMs: 1000,
      totalCards: 0,
      totalTokensSeen: 0,
      totalLinesSeen: 0,
      totalLookupCount: 0,
      totalLookupHits: 0,
      totalYomitanLookupCount: 0,
      episodeCount: 1,
      lastWatchedMs: 1,
    },
    episodes: [],
    anilistEntries: [],
  };
}

function findButton(container: Element, label: string): HTMLElement {
  const match = [...container.querySelectorAll('button')].find((button) =>
    (button.textContent ?? '').includes(label),
  );
  assert.ok(match, `expected a "${label}" button`);
  return match as unknown as HTMLElement;
}

test('AnimeTab refetches the library after the AniList entry is relinked', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    getAnimeDetail: apiClient.getAnimeDetail,
    getAnimeWords: apiClient.getAnimeWords,
    getAnimeRollups: apiClient.getAnimeRollups,
    getAnimeKnownWordsSummary: apiClient.getAnimeKnownWordsSummary,
    searchAnilist: apiClient.searchAnilist,
    reassignAnimeAnilist: apiClient.reassignAnimeAnilist,
  };

  // The library grid keys its cover URL off anilistId, so a stale list keeps
  // pointing at the previous entry's art.
  let anilistId: number | null = 14813;
  let libraryFetches = 0;

  apiClient.getAnimeLibrary = (async () => {
    libraryFetches += 1;
    return [libraryItem(anilistId)];
  }) as typeof apiClient.getAnimeLibrary;
  apiClient.getAnimeDetail = (async () => detailData(anilistId)) as typeof apiClient.getAnimeDetail;
  apiClient.getAnimeWords = (async () => []) as typeof apiClient.getAnimeWords;
  apiClient.getAnimeRollups = (async () => []) as typeof apiClient.getAnimeRollups;
  apiClient.getAnimeKnownWordsSummary = (async () => ({
    totalUniqueWords: 0,
    knownWordCount: 0,
  })) as typeof apiClient.getAnimeKnownWordsSummary;
  apiClient.searchAnilist = (async () => [
    {
      id: 108489,
      episodes: 12,
      season: null,
      seasonYear: null,
      description: null,
      coverImage: null,
      title: { romaji: 'Test Anime Kan', english: null, native: null },
    },
  ]) as typeof apiClient.searchAnilist;
  apiClient.reassignAnimeAnilist = (async (_animeId: number, info: { anilistId: number }) => {
    anilistId = info.anilistId;
  }) as typeof apiClient.reassignAnimeAnilist;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(<AnimeTab />);
    });
    assert.equal(libraryFetches, 1);
    assert.match(
      container.querySelector('img')?.getAttribute('src') ?? '',
      /\/api\/stats\/anime\/7\/cover\?coverRetry=14813/,
    );

    await act(async () => {
      findButton(container, 'Test Anime Season 2').click();
    });
    await act(async () => {
      findButton(container, 'Change AniList Entry').click();
    });
    await act(async () => {
      findButton(container, 'Test Anime Kan').click();
    });

    await act(async () => {
      findButton(container, 'Back to Library').click();
    });

    assert.equal(libraryFetches, 2);
    assert.match(
      container.querySelector('img')?.getAttribute('src') ?? '',
      /\/api\/stats\/anime\/7\/cover\?coverRetry=108489/,
    );

    await act(async () => {
      root.unmount();
    });
  } finally {
    Object.assign(apiClient, original);
    uninstallDom();
  }
});

test('AnimeTab watch time follows the displayed kind filter', async () => {
  const uninstallDom = installDom();
  const original = apiClient.getAnimeLibrary;
  apiClient.getAnimeLibrary = async () => [
    { ...libraryItem(42), totalActiveMs: 3600000 },
    {
      ...libraryItem(null),
      animeId: 8,
      canonicalTitle: 'Drama',
      mediaKind: 'live_action',
      tmdbId: 12,
      tmdbType: 'tv',
      totalActiveMs: 7200000,
    },
  ];
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => {
      root.render(<AnimeTab />);
    });
    assert.match(container.textContent ?? '', /2 titles · 3h/);
    await act(async () => {
      findButton(container, 'Live Action').click();
    });
    assert.match(container.textContent ?? '', /1 title · 2h/);
    assert.equal(findButton(container, 'Live Action').getAttribute('aria-pressed'), 'true');
  } finally {
    await act(async () => {
      root.unmount();
    });
    apiClient.getAnimeLibrary = original;
    uninstallDom();
  }
});

test('Library kind filter separates YouTube channels from anime and updates totals', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    getAnimeMergeRecommendations: apiClient.getAnimeMergeRecommendations,
  };
  apiClient.getAnimeLibrary = async () => [
    { ...libraryItem(null), totalActiveMs: 60_000 },
    {
      ...libraryItem(null),
      animeId: 8,
      mediaKind: 'youtube',
      canonicalTitle: 'Language Channel',
      totalActiveMs: 120_000,
    },
  ];
  apiClient.getAnimeMergeRecommendations = async () => ({ recommendations: [] });
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => {
      root.render(<AnimeTab />);
    });
    assert.match(container.textContent ?? '', /Test Anime Season 2/);
    assert.match(container.textContent ?? '', /Language Channel/);
    assert.match(container.textContent ?? '', /2 titles · 3m/);
    await act(async () => {
      findButton(container, 'YouTube').click();
    });
    assert.doesNotMatch(container.textContent ?? '', /Test Anime Season 2/);
    assert.match(container.textContent ?? '', /Language Channel/);
    assert.match(container.textContent ?? '', /1 channel · 2m/);
    assert.equal(findButton(container, 'YouTube').getAttribute('aria-pressed'), 'true');
    await act(async () => {
      findButton(container, 'Anime').click();
    });
    assert.match(container.textContent ?? '', /Test Anime Season 2/);
    assert.doesNotMatch(container.textContent ?? '', /Language Channel/);
    await act(async () => {
      findButton(container, 'All Titles').click();
    });
    assert.match(container.textContent ?? '', /Language Channel/);
  } finally {
    await act(async () => {
      root.unmount();
    });
    Object.assign(apiClient, original);
    uninstallDom();
  }
});
