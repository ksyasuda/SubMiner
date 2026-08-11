import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Database } from '../sqlite.js';
import type { DatabaseSync } from '../sqlite.js';
import { applyPragmas, ensureSchema } from '../storage.js';
import { mergeAnimeRecords, moveVideoToAnime } from '../anime-merge.js';
import { resolveAnimeAnilistConflict } from '../anime-season-repair.js';

const BASE_MS = 1_700_000_000_000;

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-anime-merge-test-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  const dir = path.dirname(dbPath);
  if (!fs.existsSync(dir)) return;
  fs.rmSync(dir, { recursive: true, force: true });
}

function withDb(work: (db: DatabaseSync) => void): void {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  try {
    applyPragmas(db);
    ensureSchema(db);
    work(db);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
}

interface AnimeSeed {
  animeId: number;
  key: string;
  title: string;
  anilistId?: number | null;
  titleRomaji?: string | null;
}

function insertAnime(db: DatabaseSync, seed: AnimeSeed): void {
  db.prepare(
    `INSERT INTO imm_anime(anime_id, normalized_title_key, canonical_title, anilist_id, title_romaji, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
  ).run(
    seed.animeId,
    seed.key,
    seed.title,
    seed.anilistId ?? null,
    seed.titleRomaji ?? null,
    BASE_MS,
    BASE_MS,
  );
}

interface EpisodeSeed {
  videoId: number;
  animeId: number;
  season?: number | null;
  episode?: number;
  activeMs?: number;
  cards?: number;
}

/** One episode with one ended session, so lifetime rebuilds have something to sum. */
function insertEpisode(db: DatabaseSync, seed: EpisodeSeed): void {
  const activeMs = seed.activeMs ?? 1000;
  const cards = seed.cards ?? 1;
  db.prepare(
    `INSERT INTO imm_videos(video_id, video_key, anime_id, canonical_title, source_type, parsed_title, parsed_season, parsed_episode, watched, duration_ms, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, ?, ?, 1, 'Show', ?, ?, 1, 1440000, ?, ?)`,
  ).run(
    seed.videoId,
    `local:/tmp/show-${seed.videoId}.mkv`,
    seed.animeId,
    `Show ${seed.videoId}`,
    seed.season ?? null,
    seed.episode ?? seed.videoId,
    BASE_MS,
    BASE_MS,
  );
  db.prepare(
    `INSERT INTO imm_sessions(session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, active_watched_ms, cards_mined, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, ?, ?, ?, 2, ?, ?, ?, ?)`,
  ).run(
    seed.videoId,
    `session-${seed.videoId}`,
    seed.videoId,
    String(BASE_MS),
    String(BASE_MS + activeMs),
    activeMs,
    cards,
    BASE_MS,
    BASE_MS,
  );
  db.prepare(
    `INSERT INTO imm_subtitle_lines(session_id, video_id, anime_id, line_index, text, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, ?, 1, ?, ?, ?)`,
  ).run(seed.videoId, seed.videoId, seed.animeId, `line ${seed.videoId}`, BASE_MS, BASE_MS);
}

function animeIds(db: DatabaseSync): number[] {
  return (
    db.prepare('SELECT anime_id AS id FROM imm_anime ORDER BY anime_id').all() as Array<{
      id: number;
    }>
  ).map((row) => row.id);
}

function videoAnimeId(db: DatabaseSync, videoId: number): number | null {
  return (
    db.prepare('SELECT anime_id AS id FROM imm_videos WHERE video_id = ?').get(videoId) as {
      id: number | null;
    }
  ).id;
}

function lineAnimeIds(db: DatabaseSync, animeId: number): number {
  return Number(
    (
      db
        .prepare('SELECT COUNT(*) AS total FROM imm_subtitle_lines WHERE anime_id = ?')
        .get(animeId) as { total: number }
    ).total,
  );
}

test('mergeAnimeRecords folds episodes, lines and lifetime totals into the target', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show' });
    insertAnime(db, { animeId: 2, key: 'show season 1', title: 'Show Season 1', anilistId: 555 });
    insertEpisode(db, { videoId: 1, animeId: 1, activeMs: 1000, cards: 1 });
    insertEpisode(db, { videoId: 2, animeId: 2, season: 1, activeMs: 2000, cards: 3 });

    const summary = mergeAnimeRecords(db, 1, [2]);

    assert.equal(summary.survivingAnimeId, 1);
    assert.deepEqual(summary.mergedAnimeIds, [2]);
    assert.equal(summary.movedVideos, 1);
    assert.deepEqual(animeIds(db), [1]);
    assert.equal(videoAnimeId(db, 2), 1);
    assert.equal(lineAnimeIds(db, 1), 2);

    const lifetime = db
      .prepare(
        'SELECT total_active_ms AS activeMs, total_cards AS cards, episodes_started AS episodes FROM imm_lifetime_anime WHERE anime_id = 1',
      )
      .get() as { activeMs: number; cards: number; episodes: number };
    assert.equal(lifetime.activeMs, 3000);
    assert.equal(lifetime.cards, 4);
    assert.equal(lifetime.episodes, 2);
  });
});

test('mergeAnimeRecords repoints subtitle lines recorded before the anime link landed', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show' });
    insertAnime(db, { animeId: 2, key: 'show season 1', title: 'Show Season 1' });
    insertEpisode(db, { videoId: 1, animeId: 1 });
    insertEpisode(db, { videoId: 2, animeId: 2, season: 1 });
    // Lines are written with the video's anime_id at the time, which is NULL
    // until the async title parse assigns one.
    db.prepare(
      `INSERT INTO imm_subtitle_lines(session_id, video_id, anime_id, line_index, text, CREATED_DATE, LAST_UPDATE_DATE)
       VALUES (2, 2, NULL, 2, 'unlinked line', ?, ?)`,
    ).run(BASE_MS, BASE_MS);

    mergeAnimeRecords(db, 1, [2]);

    assert.equal(lineAnimeIds(db, 1), 3);
    const orphaned = Number(
      (
        db
          .prepare('SELECT COUNT(*) AS total FROM imm_subtitle_lines WHERE anime_id IS NULL')
          .get() as { total: number }
      ).total,
    );
    assert.equal(orphaned, 0);
  });
});

test('mergeAnimeRecords inherits metadata the target is missing without clobbering its own', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show', titleRomaji: 'Shou' });
    insertAnime(db, {
      animeId: 2,
      key: 'show season 1',
      title: 'Show Season 1',
      anilistId: 555,
      titleRomaji: 'Show Romaji',
    });
    insertEpisode(db, { videoId: 1, animeId: 1 });
    insertEpisode(db, { videoId: 2, animeId: 2, season: 1 });

    mergeAnimeRecords(db, 1, [2]);

    const row = db
      .prepare(
        'SELECT canonical_title AS title, anilist_id AS anilistId, title_romaji AS romaji FROM imm_anime WHERE anime_id = 1',
      )
      .get() as { title: string; anilistId: number | null; romaji: string | null };
    assert.equal(row.title, 'Show');
    // anilist_id is UNIQUE, so inheriting it proves the source row was gone first.
    assert.equal(row.anilistId, 555);
    assert.equal(row.romaji, 'Shou');
  });
});

test('mergeAnimeRecords ignores unknown targets and self-merges', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show' });
    insertEpisode(db, { videoId: 1, animeId: 1 });

    assert.deepEqual(mergeAnimeRecords(db, 99, [1]).mergedAnimeIds, []);
    assert.deepEqual(mergeAnimeRecords(db, 1, [1]).mergedAnimeIds, []);
    assert.deepEqual(animeIds(db), [1]);
    assert.equal(videoAnimeId(db, 1), 1);
  });
});

test('moveVideoToAnime moves one episode and prunes the emptied entry', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show' });
    insertAnime(db, { animeId: 2, key: 'stray', title: 'Stray Episode Title', anilistId: 777 });
    insertEpisode(db, { videoId: 1, animeId: 1 });
    insertEpisode(db, { videoId: 2, animeId: 2, activeMs: 5000, cards: 2 });

    const summary = moveVideoToAnime(db, 2, 1);

    assert.equal(summary.targetAnimeId, 1);
    assert.equal(summary.previousAnimeId, 2);
    assert.equal(summary.removedPreviousAnime, true);
    assert.deepEqual(animeIds(db), [1]);
    assert.equal(videoAnimeId(db, 2), 1);
    assert.equal(lineAnimeIds(db, 1), 2);
    const lifetime = db
      .prepare('SELECT total_active_ms AS activeMs FROM imm_lifetime_anime WHERE anime_id = 1')
      .get() as { activeMs: number };
    assert.equal(lifetime.activeMs, 6000);
    // The stray entry's AniList link is dropped, not inherited: a move makes no
    // claim that the two entries are the same show.
    const target = db
      .prepare('SELECT anilist_id AS anilistId FROM imm_anime WHERE anime_id = 1')
      .get() as { anilistId: number | null };
    assert.equal(target.anilistId, null);
  });
});

test('moveVideoToAnime is a no-op when the episode is already in the target entry', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show' });
    insertEpisode(db, { videoId: 1, animeId: 1 });

    const summary = moveVideoToAnime(db, 1, 1);

    assert.equal(summary.targetAnimeId, 1);
    assert.equal(summary.previousAnimeId, 1);
    assert.equal(summary.removedPreviousAnime, false);
    assert.deepEqual(animeIds(db), [1]);
    assert.equal(videoAnimeId(db, 1), 1);
  });
});

test('moveVideoToAnime keeps the source entry when other episodes remain', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show' });
    insertAnime(db, { animeId: 2, key: 'other', title: 'Other' });
    insertEpisode(db, { videoId: 1, animeId: 2 });
    insertEpisode(db, { videoId: 2, animeId: 2 });

    const summary = moveVideoToAnime(db, 2, 1);

    assert.equal(summary.removedPreviousAnime, false);
    assert.deepEqual(animeIds(db), [1, 2]);
    assert.equal(videoAnimeId(db, 1), 2);
    assert.equal(videoAnimeId(db, 2), 1);
  });
});

test('moveVideoToAnime rejects unknown episodes and targets', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show' });
    insertEpisode(db, { videoId: 1, animeId: 1 });

    assert.throws(() => moveVideoToAnime(db, 99, 1));
    assert.throws(() => moveVideoToAnime(db, 1, 99));
    assert.equal(videoAnimeId(db, 1), 1);
  });
});

test('resolveAnimeAnilistConflict folds a seasonless duplicate into the entry that owns the id', () => {
  withDb((db) => {
    // Same show, split because one release tagged S01 and the other did not.
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show', anilistId: 163132 });
    insertAnime(db, { animeId: 2, key: 'show season 1', title: 'Show Season 1' });
    insertEpisode(db, { videoId: 1, animeId: 1 });
    insertEpisode(db, { videoId: 2, animeId: 2, season: 1 });

    const summary = resolveAnimeAnilistConflict(db, 2, 163132);

    assert.equal(summary.survivingAnimeId, 1);
    assert.equal(summary.movedVideos, 1);
    assert.equal(summary.deletedAnimeRows, 1);
    assert.deepEqual(animeIds(db), [1]);
    assert.equal(videoAnimeId(db, 2), 1);
  });
});

test('resolveAnimeAnilistConflict keeps the target entry when the user drove the change', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show', anilistId: 163132 });
    insertAnime(db, { animeId: 2, key: 'show season 1', title: 'Show Season 1' });
    insertEpisode(db, { videoId: 1, animeId: 1 });
    insertEpisode(db, { videoId: 2, animeId: 2, season: 1 });

    const summary = resolveAnimeAnilistConflict(db, 2, 163132, { survivor: 'target' });

    assert.equal(summary.survivingAnimeId, 2);
    assert.deepEqual(animeIds(db), [2]);
    assert.equal(videoAnimeId(db, 1), 2);
    const row = db.prepare('SELECT anilist_id AS id FROM imm_anime WHERE anime_id = 2').get() as {
      id: number | null;
    };
    assert.equal(row.id, 163132);
  });
});

test('resolveAnimeAnilistConflict falls back to season redistribution for multi-season rows', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show', anilistId: 163132 });
    insertAnime(db, { animeId: 2, key: 'show season 1', title: 'Show Season 1' });
    insertEpisode(db, { videoId: 1, animeId: 1, season: 1 });
    insertEpisode(db, { videoId: 2, animeId: 1, season: 2 });
    insertEpisode(db, { videoId: 3, animeId: 2, season: 1 });

    resolveAnimeAnilistConflict(db, 2, 163132);

    // The mixed row is split by season instead of being poured onto one card.
    const titles = (
      db.prepare('SELECT canonical_title AS title FROM imm_anime ORDER BY title').all() as Array<{
        title: string;
      }>
    ).map((row) => row.title);
    assert.deepEqual(titles, ['Show Season 1', 'Show Season 2']);
    assert.equal(videoAnimeId(db, 1), 2);
    assert.equal(videoAnimeId(db, 3), 2);
    assert.notEqual(videoAnimeId(db, 2), 2);
  });
});

test('resolveAnimeAnilistConflict leaves an entry that already links elsewhere alone', () => {
  withDb((db) => {
    insertAnime(db, { animeId: 1, key: 'show', title: 'Show', anilistId: 163132 });
    insertAnime(db, { animeId: 2, key: 'show s2', title: 'Show Season 2', anilistId: 999 });
    insertEpisode(db, { videoId: 1, animeId: 1, season: 1 });
    insertEpisode(db, { videoId: 2, animeId: 2, season: 2 });

    resolveAnimeAnilistConflict(db, 2, 163132);

    assert.equal(videoAnimeId(db, 2), 2);
    assert.ok(animeIds(db).includes(2));
  });
});
