import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { AnimeHeader } from './AnimeHeader';
import { confirmAnimeDelete, setDeleteConfirmPresenter } from '../../lib/delete-confirm';
import type { AnimeDetailData } from '../../types/stats';

const DETAIL: AnimeDetailData['detail'] = {
  mediaKind: 'anime',
  animeId: 3,
  canonicalTitle: 'Project Radio Noise Season 2',
  anilistId: 20661,
  tmdbId: null,
  tmdbType: null,
  titleRomaji: 'Toaru Kagaku no Railgun S',
  titleEnglish: 'A Certain Scientific Railgun S',
  titleNative: null,
  description: null,
  totalSessions: 1,
  totalActiveMs: 960_000,
  totalCards: 0,
  totalTokensSeen: 1_655,
  totalLinesSeen: 300,
  totalLookupCount: 0,
  totalLookupHits: 0,
  totalYomitanLookupCount: 0,
  episodeCount: 1,
  lastWatchedMs: 1_700_000_000_000,
};

test('AnimeHeader offers a delete-entry action alongside the AniList actions', () => {
  const markup = renderToStaticMarkup(
    <AnimeHeader detail={DETAIL} anilistEntries={[]} onDeleteAnime={() => {}} />,
  );

  assert.match(markup, /Delete Entry/);
  assert.match(markup, /Delete this title and every session and stat recorded for it/);
});

test('AnimeHeader hides the delete action when no handler is wired', () => {
  const markup = renderToStaticMarkup(<AnimeHeader detail={DETAIL} anilistEntries={[]} />);

  assert.doesNotMatch(markup, /Delete Entry/);
});

test('AnimeHeader shows a pending state while the entry is being deleted', () => {
  const markup = renderToStaticMarkup(
    <AnimeHeader detail={DETAIL} anilistEntries={[]} onDeleteAnime={() => {}} isDeletingAnime />,
  );

  assert.match(markup, /disabled=""/);
  assert.match(markup, /animate-spin/);
  assert.doesNotMatch(markup, /Delete Entry/);
});

test('confirmAnimeDelete spells out how much data the entry deletion removes', async () => {
  const seen: string[] = [];
  const restore = setDeleteConfirmPresenter((message) => {
    seen.push(message);
    return false;
  });

  try {
    await confirmAnimeDelete('Railgun Season 2', 1);
    await confirmAnimeDelete('Railgun Season 2', 3);
  } finally {
    restore();
  }

  assert.match(seen[0] ?? '', /"Railgun Season 2"/);
  assert.match(seen[0] ?? '', /1 episode\b/);
  assert.match(seen[1] ?? '', /3 episodes/);
  assert.match(seen[0] ?? '', /every session and stat/);
});

test('AnimeHeader shows TMDB actions for a live-action entry and hides AniList links', () => {
  const markup = renderToStaticMarkup(
    <AnimeHeader
      detail={{
        ...DETAIL,
        anilistId: null,
        mediaKind: 'live_action',
        tmdbId: 61222,
        tmdbType: 'tv',
        titleRomaji: null,
        titleEnglish: 'Hanzawa Naoki',
        titleNative: '半沢直樹',
        description: 'A banker fights back.',
      }}
      anilistEntries={[]}
      onChangeAnilist={() => {}}
      onChangeTmdb={() => {}}
    />,
  );

  assert.match(markup, /https:\/\/www\.themoviedb\.org\/tv\/61222/);
  assert.match(markup, /Change TMDB Title/);
  assert.match(markup, /Live action/);
  assert.match(markup, /A banker fights back\./);
  assert.doesNotMatch(markup, /anilist\.co/);
});

test('AnimeHeader offers to link an unlinked anime entry to TMDB', () => {
  const markup = renderToStaticMarkup(
    <AnimeHeader
      detail={{ ...DETAIL, anilistId: null }}
      anilistEntries={[]}
      onChangeTmdb={() => {}}
    />,
  );

  assert.match(markup, /Link to TMDB/);
  assert.doesNotMatch(markup, /themoviedb\.org/);
});

test('YouTube channel headers show videos and omit all AniList controls', () => {
  const markup = renderToStaticMarkup(
    <AnimeHeader
      detail={{ ...DETAIL, mediaKind: 'youtube' }}
      anilistEntries={[{ anilistId: 20661, titleRomaji: null, titleEnglish: null, season: null }]}
      onChangeAnilist={() => {}}
    />,
  );
  assert.match(markup, /YouTube channel/);
  assert.match(markup, /video/);
  assert.doesNotMatch(markup, /AniList|anilist\.co|episode/);
});
