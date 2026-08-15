import assert from 'node:assert/strict';
import test from 'node:test';
import { rebuildLifetimeSummaries, repairLifetimeSummariesFromMedia } from '../lifetime.js';
import {
  BASE_MS,
  DAY_MS,
  cleanRow,
  createDb,
  seedAnime,
  seedEndedSession,
  seedVideo,
  snapshotAnime,
  snapshotGlobal,
  snapshotMedia,
} from './lifetime-test-fixtures.js';

test('repair after a video moves between anime matches a full rebuild', () => {
  const db = createDb();
  try {
    const animeA = seedAnime(db, 'Move Source', 2);
    const animeB = seedAnime(db, 'Move Target', 2);
    const movedVideo = seedVideo(db, animeA, 'moved-ep', { watched: true });
    const stayingVideo = seedVideo(db, animeA, 'staying-ep');
    const targetVideo = seedVideo(db, animeB, 'target-ep', { watched: true });
    seedEndedSession(db, movedVideo, BASE_MS, { activeMs: 60_000, cards: 2 });
    seedEndedSession(db, stayingVideo, BASE_MS + DAY_MS, { activeMs: 30_000 });
    seedEndedSession(db, targetVideo, BASE_MS + 2 * DAY_MS, { activeMs: 45_000, cards: 1 });
    rebuildLifetimeSummaries(db);

    // Simulate a library merge reassigning the episode to the other anime.
    db.prepare('UPDATE imm_videos SET anime_id = ? WHERE video_id = ?').run(animeB, movedVideo);
    repairLifetimeSummariesFromMedia(db);

    const repairedGlobal = snapshotGlobal(db);
    const repairedMedia = snapshotMedia(db);
    const repairedAnime = snapshotAnime(db);

    rebuildLifetimeSummaries(db);
    assert.deepEqual(repairedGlobal, snapshotGlobal(db));
    assert.deepEqual(repairedMedia, snapshotMedia(db));
    assert.deepEqual(repairedAnime, snapshotAnime(db));
  } finally {
    db.close();
  }
});

test('repair preserves lifetime history from pruned sessions where a rebuild would not', () => {
  const db = createDb();
  try {
    const animeId = seedAnime(db, 'Repair Anime', null);
    const videoId = seedVideo(db, animeId, 'repair-ep');
    const prunedSessionId = seedEndedSession(db, videoId, BASE_MS, {
      activeMs: 90_000,
      cards: 3,
    });
    seedEndedSession(db, videoId, BASE_MS + DAY_MS, { activeMs: 30_000, cards: 1 });
    rebuildLifetimeSummaries(db);

    db.prepare('DELETE FROM imm_sessions WHERE session_id = ?').run(prunedSessionId);
    repairLifetimeSummariesFromMedia(db);

    const globalRow = snapshotGlobal(db);
    assert.equal(globalRow.total_sessions, 2, 'repair keeps the pruned session contribution');
    assert.equal(globalRow.total_active_ms, 120_000);
    assert.equal(globalRow.total_cards, 4);
    assert.equal(globalRow.active_days, 2, 'repair never subtracts active days');

    const animeRow = db
      .prepare('SELECT total_sessions FROM imm_lifetime_anime WHERE anime_id = ?')
      .get(animeId);
    assert.equal(cleanRow<{ total_sessions: number }>(animeRow).total_sessions, 2);
  } finally {
    db.close();
  }
});

test('repair leaves a caller-owned transaction intact when its begin fails', () => {
  const db = createDb();
  try {
    db.exec('BEGIN');
    const animeId = seedAnime(db, 'Caller Transaction', null);

    assert.throws(() => repairLifetimeSummariesFromMedia(db), /transaction/i);
    assert.ok(
      db.prepare('SELECT 1 FROM imm_anime WHERE anime_id = ?').get(animeId),
      'the repair did not roll back the caller transaction',
    );

    db.exec('ROLLBACK');
    assert.equal(db.prepare('SELECT 1 FROM imm_anime WHERE anime_id = ?').get(animeId), undefined);
  } finally {
    db.close();
  }
});
