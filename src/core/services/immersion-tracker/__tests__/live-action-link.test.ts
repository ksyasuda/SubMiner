import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Database } from '../sqlite.js';
import type { DatabaseSync } from '../sqlite.js';
import { applyPragmas, ensureSchema, getOrCreateAnimeRecord } from '../storage.js';
import { repairLegacySeasonlessAnimeRows } from '../anime-season-repair.js';
import { mergeAnimeRecords, mergeAnimeRecordsInTransaction } from '../anime-merge.js';
import { getVideoTmdbLink, linkAnimeToTmdbTitle } from '../live-action-link.js';
import { getAnimeCoverArt, getCoverArt } from '../query-library.js';
import { clearAnimeCoverArt, upsertCoverArt } from '../query-maintenance.js';

const BASE_MS = 1_700_000_000_000;

function withDb(work: (db: DatabaseSync) => void): void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-live-action-link-'));
  const db = new Database(path.join(dir, 'immersion.sqlite'));
  try {
    applyPragmas(db);
    ensureSchema(db);
    work(db);
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function insertAnime(
  db: DatabaseSync,
  animeId: number,
  title: string,
  anilistId: number | null = null,
) {
  db.prepare(
    `INSERT INTO imm_anime(anime_id, normalized_title_key, canonical_title, anilist_id, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, ?, ?, ?, ?)`,
  ).run(animeId, title.toLowerCase(), title, anilistId, BASE_MS, BASE_MS);
}

function insertEpisode(db: DatabaseSync, videoId: number, animeId: number, season: number | null) {
  db.prepare(
    `INSERT INTO imm_videos(video_id, video_key, anime_id, canonical_title, source_type, parsed_title, parsed_season, parsed_episode, watched, duration_ms, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, ?, ?, 1, 'Hanzawa Naoki', ?, ?, 1, 1440000, ?, ?)`,
  ).run(
    videoId,
    `local:/tmp/${videoId}.mkv`,
    animeId,
    `Ep ${videoId}`,
    season,
    videoId,
    BASE_MS,
    BASE_MS,
  );
  db.prepare(
    `INSERT INTO imm_lifetime_media(video_id, total_sessions, total_active_ms, total_cards, completed, first_watched_ms, last_watched_ms, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, 1, 1000, 0, 1, ?, ?, ?, ?)`,
  ).run(videoId, String(BASE_MS), String(BASE_MS + 1000), BASE_MS, BASE_MS);
}

interface AnimeRowView {
  mediaKind: string;
  tmdbId: number | null;
  tmdbType: string | null;
  anilistId: number | null;
  titleEnglish: string | null;
  titleNative: string | null;
  description: string | null;
}

// Copies the selected columns so the driver's row metadata does not leak into
// deep-equality assertions.
function animeRow(db: DatabaseSync, animeId: number): AnimeRowView | undefined {
  const row = db
    .prepare(
      `SELECT media_kind AS mediaKind, tmdb_id AS tmdbId, tmdb_type AS tmdbType, anilist_id AS anilistId,
              title_english AS titleEnglish, title_native AS titleNative, description
       FROM imm_anime WHERE anime_id = ?`,
    )
    .get(animeId) as AnimeRowView | undefined;
  if (!row) return undefined;
  const { mediaKind, tmdbId, tmdbType, anilistId, titleEnglish, titleNative, description } = row;
  return { mediaKind, tmdbId, tmdbType, anilistId, titleEnglish, titleNative, description };
}

function animeCount(db: DatabaseSync): number {
  return (db.prepare('SELECT COUNT(*) AS n FROM imm_anime').get() as { n: number }).n;
}

function videoOwner(db: DatabaseSync, videoId: number): number | null {
  return (
    db.prepare('SELECT anime_id AS animeId FROM imm_videos WHERE video_id = ?').get(videoId) as {
      animeId: number | null;
    }
  ).animeId;
}

const HANZAWA = {
  tmdbId: 61222,
  tmdbType: 'tv' as const,
  titleEnglish: 'Hanzawa Naoki',
  titleNative: '半沢直樹',
  description: 'A banker fights back.',
  episodesTotal: 10,
};

test('a manual link overwrites metadata, drops the AniList link, and folds other holders in', () => {
  withDb((db) => {
    insertAnime(db, 1, 'Hanzawa Naoki', 4242);
    insertAnime(db, 2, 'Hanzawa Naoki Season 2');
    insertEpisode(db, 1, 1, 1);
    insertEpisode(db, 2, 2, 2);
    db.prepare(
      `UPDATE imm_anime SET media_kind = 'live_action', tmdb_id = ?, tmdb_type = 'tv', description = 'old' WHERE anime_id = 2`,
    ).run(HANZAWA.tmdbId);

    const result = linkAnimeToTmdbTitle(db, 1, HANZAWA, { mode: 'manual' });

    assert.deepEqual(result, { animeId: 1, mergedAnimeIds: [2] });
    assert.deepEqual(animeRow(db, 1), {
      mediaKind: 'live_action',
      tmdbId: 61222,
      tmdbType: 'tv',
      anilistId: null,
      titleEnglish: 'Hanzawa Naoki',
      titleNative: '半沢直樹',
      description: 'A banker fights back.',
    });
    assert.equal(animeRow(db, 2), undefined);
    assert.equal(animeCount(db), 1);
    assert.equal(videoOwner(db, 2), 1);
    assert.deepEqual(getVideoTmdbLink(db, 2), { animeId: 1, tmdbId: 61222, tmdbType: 'tv' });
  });
});

test('an automatic link joins the entry that already owns the title and only fills gaps', () => {
  withDb((db) => {
    insertAnime(db, 1, 'Hanzawa Naoki');
    insertEpisode(db, 1, 1, 1);
    linkAnimeToTmdbTitle(db, 1, { ...HANZAWA, description: 'kept' }, { mode: 'manual' });
    const newcomer = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Hanzawa Naoki',
      canonicalTitle: 'Hanzawa Naoki',
      seasonScope: 2,
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    insertEpisode(db, 2, newcomer, 2);

    const result = linkAnimeToTmdbTitle(db, newcomer, HANZAWA, { mode: 'auto' });

    assert.deepEqual(result, { animeId: 1, mergedAnimeIds: [newcomer] });
    assert.equal(animeRow(db, 1)?.description, 'kept');
    assert.equal(animeCount(db), 1);
    assert.equal(videoOwner(db, 2), 1);
    // The merged-away season title is remembered, so the next episode of that
    // season lands on the survivor without a detour through a new row.
    const again = getOrCreateAnimeRecord(db, {
      parsedTitle: 'Hanzawa Naoki',
      canonicalTitle: 'Hanzawa Naoki',
      seasonScope: 2,
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    assert.equal(again, 1);
  });
});

test('startup season repair leaves multi-season live-action entries alone', () => {
  withDb((db) => {
    insertAnime(db, 1, 'Hanzawa Naoki');
    insertEpisode(db, 1, 1, 1);
    insertEpisode(db, 2, 1, 2);
    linkAnimeToTmdbTitle(db, 1, HANZAWA, { mode: 'manual' });

    repairLegacySeasonlessAnimeRows(db);

    assert.equal(animeCount(db), 1);
    assert.equal(videoOwner(db, 1), 1);
    assert.equal(videoOwner(db, 2), 1);
    assert.equal(getVideoTmdbLink(db, 2)?.animeId, 1);
  });
});

test('getVideoTmdbLink is null for anime entries and unlinked videos', () => {
  withDb((db) => {
    insertAnime(db, 1, 'Some Anime', 77);
    insertEpisode(db, 1, 1, 1);
    assert.equal(getVideoTmdbLink(db, 1), null);
    assert.equal(getVideoTmdbLink(db, 99), null);
  });
});

test('clearAnimeCoverArt drops every episode cover of the entry and its orphaned blob', () => {
  withDb((db) => {
    insertAnime(db, 1, 'Hanzawa Naoki');
    insertAnime(db, 2, 'Other Show');
    insertEpisode(db, 1, 1, 1);
    insertEpisode(db, 2, 1, 1);
    insertEpisode(db, 3, 2, 1);
    const shared = Buffer.from([1, 2, 3]);
    for (const videoId of [1, 2]) {
      upsertCoverArt(db, videoId, {
        anilistId: 4242,
        coverUrl: 'https://images.test/a.jpg',
        coverBlob: shared,
        titleRomaji: null,
        titleEnglish: null,
        episodesTotal: null,
      });
    }
    upsertCoverArt(db, 3, {
      anilistId: 99,
      coverUrl: 'https://images.test/b.jpg',
      coverBlob: Buffer.from([9]),
      titleRomaji: null,
      titleEnglish: null,
      episodesTotal: null,
    });

    clearAnimeCoverArt(db, 1);

    assert.equal(getAnimeCoverArt(db, 1), null);
    assert.equal(getCoverArt(db, 3)?.coverBlob?.length, 1);
    const blobs = (
      db.prepare('SELECT COUNT(*) AS n FROM imm_cover_art_blobs').get() as { n: number }
    ).n;
    assert.equal(blobs, 1);
  });
});

for (const targetId of [1, 2, 3]) {
  test(`merge rejects mixed providers before moving any source into entry ${targetId}`, () => {
    withDb((db) => {
      insertAnime(db, 1, 'Anime', 77);
      insertAnime(db, 2, 'Drama');
      insertAnime(db, 3, 'Unlinked');
      insertEpisode(db, 1, 1, 1);
      insertEpisode(db, 2, 2, 1);
      db.exec(
        "UPDATE imm_anime SET media_kind = 'live_action', tmdb_id = 12, tmdb_type = 'tv' WHERE anime_id = 2",
      );
      for (const merge of [mergeAnimeRecords, mergeAnimeRecordsInTransaction]) {
        assert.throws(
          () => merge(db, targetId, [3, 1, 2]),
          /AniList-linked and TMDB-linked library entries cannot be merged/,
        );
        assert.equal(animeCount(db), 3);
        assert.equal(videoOwner(db, 1), 1);
        assert.equal(videoOwner(db, 2), 2);
      }
    });
  });
}

for (const mode of ['manual', 'auto'] as const) {
  test(`TMDB ${mode} linking rolls back the merge when the survivor update fails`, () => {
    withDb((db) => {
      insertAnime(db, 1, 'New entry');
      insertAnime(db, 2, 'Existing entry');
      insertEpisode(db, 1, 1, 1);
      insertEpisode(db, 2, 2, 2);
      db.prepare(
        "UPDATE imm_anime SET tmdb_id = ?, tmdb_type = 'tv', media_kind = 'live_action' WHERE anime_id = 2",
      ).run(HANZAWA.tmdbId);
      db.exec(`CREATE TRIGGER reject_link BEFORE UPDATE ON imm_anime
        WHEN NEW.description = 'A banker fights back.'
        BEGIN SELECT RAISE(ABORT, 'rejected survivor update'); END`);
      assert.throws(
        () => linkAnimeToTmdbTitle(db, 1, HANZAWA, { mode }),
        /rejected survivor update/,
      );
      assert.equal(animeCount(db), 2);
      assert.equal(videoOwner(db, 1), 1);
      assert.equal(videoOwner(db, 2), 2);
      assert.equal(animeRow(db, 1)?.tmdbId, null);
      assert.equal(animeRow(db, 2)?.tmdbId, HANZAWA.tmdbId);
    });
  });
}

for (const mode of ['manual', 'auto'] as const) {
  test(`${mode} TMDB linking refreshes completion totals without merging records`, () => {
    withDb((db) => {
      insertAnime(db, 1, 'Hanzawa Naoki');
      insertEpisode(db, 1, 1, 1);
      const completed = () =>
        (
          db
            .prepare('SELECT anime_completed AS count FROM imm_lifetime_global WHERE global_id = 1')
            .get() as { count: number }
        ).count;
      assert.equal(completed(), 0);
      const result = linkAnimeToTmdbTitle(db, 1, { ...HANZAWA, episodesTotal: 1 }, { mode });
      assert.deepEqual(result.mergedAnimeIds, []);
      assert.equal(completed(), 1);
      linkAnimeToTmdbTitle(db, 1, { ...HANZAWA, episodesTotal: 2 }, { mode: 'manual' });
      assert.equal(completed(), 0);
      assert.equal(animeCount(db), 1);
    });
  });
}
