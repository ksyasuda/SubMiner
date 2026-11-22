import assert from 'node:assert/strict';
import test from 'node:test';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { apiClient } from '../../lib/api-client';
import type {
  AnimeDetailData,
  AnimeLibraryGroup,
  AnimeLibraryGroupPart,
  StatsAnimeGroupOverrideRequest,
} from '../../types/stats';
import { AnimeTab } from './AnimeTab';
import { installDom } from '../../test-utils/dom';

function part(
  animeId: number,
  title: string,
  format: string,
  startDate: number,
): AnimeLibraryGroupPart {
  return {
    mediaKind: 'anime',
    animeId,
    canonicalTitle: title,
    anilistId: animeId * 100,
    tmdbId: null,
    tmdbType: null,
    totalSessions: 1,
    totalActiveMs: 60_000,
    totalCards: 2,
    totalTokensSeen: 0,
    episodeCount: format === 'MOVIE' ? 1 : 13,
    episodesTotal: null,
    lastWatchedMs: animeId,
    format,
    startDate,
  };
}

const overlord: AnimeLibraryGroup = {
  groupKey: 20832,
  title: 'Overlord',
  mediaKind: 'anime',
  coverAnimeId: 2,
  totalSessions: 2,
  totalActiveMs: 120_000,
  totalCards: 4,
  totalTokensSeen: 0,
  episodeCount: 14,
  lastWatchedMs: 2,
  parts: [part(2, 'The Undead King', 'MOVIE', 20170225), part(1, 'Overlord II', 'TV', 20180109)],
};

function detailData(): AnimeDetailData {
  const {
    format: _format,
    startDate: _startDate,
    episodesTotal: _total,
    ...item
  } = part(1, 'Overlord II', 'TV', 20180109);
  return {
    detail: {
      ...item,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      description: null,
      totalLinesSeen: 0,
      totalLookupCount: 0,
      totalLookupHits: 0,
      totalYomitanLookupCount: 0,
    },
    episodes: [],
    anilistEntries: [],
    group: { groupKey: 20832, title: 'Overlord', partCount: 2, overrideMode: 'anilist' },
  };
}

function findButton(container: Element, label: string): HTMLElement {
  const match = [...container.querySelectorAll('button')].find((button) =>
    (button.textContent ?? '').includes(label),
  );
  assert.ok(match, `expected a "${label}" button`);
  return match as unknown as HTMLElement;
}

test('a franchise card opens its parts in release order, and a part can stand alone', async () => {
  const uninstallDom = installDom();
  const original = { ...apiClient };
  const overrides: Array<[number, StatsAnimeGroupOverrideRequest]> = [];
  let libraryFetches = 0;
  apiClient.getAnimeLibraryGroups = async () => {
    libraryFetches += 1;
    return [overlord];
  };
  apiClient.getAnimeMergeRecommendations = async () => ({ recommendations: [] });
  apiClient.getAnimeDetail = async () => detailData();
  apiClient.getAnimeWords = async () => [];
  apiClient.getAnimeRollups = async () => [];
  apiClient.getAnimeKnownWordsSummary = async () => ({ totalUniqueWords: 0, knownWordCount: 0 });
  apiClient.setAnimeGroupOverride = async (animeId, request) => {
    overrides.push([animeId, request]);
  };

  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => root.render(<AnimeTab />));
    // One card for the franchise, counted once, with the first-released part's cover.
    assert.match(container.textContent ?? '', /1 title/);
    assert.match(findButton(container, 'Overlord').textContent ?? '', /2 parts/);
    assert.match(container.querySelector('img')?.getAttribute('src') ?? '', /anime\/2\/cover/);

    await act(async () => findButton(container, 'Overlord').click());
    const rows = [...container.querySelectorAll('li button')].map((row) => row.textContent ?? '');
    assert.equal(rows.length, 2);
    assert.match(rows[0]!, /The Undead King.*Movie · 2017/);
    assert.match(rows[1]!, /Overlord II.*TV · 2018/);

    await act(async () => findButton(container, 'Overlord II').click());
    assert.match(container.textContent ?? '', /Listed under Overlord with 1 other part/);
    await act(async () => findButton(container, 'Show On Its Own').click());
    assert.deepEqual(overrides, [[1, { mode: 'standalone' }]]);
    assert.equal(libraryFetches, 2);

    await act(async () => findButton(container, 'Back to Overlord').click());
    assert.match(container.textContent ?? '', /Seasons and Movies/);
  } finally {
    await act(async () => root.unmount());
    Object.assign(apiClient, original);
    uninstallDom();
  }
});
