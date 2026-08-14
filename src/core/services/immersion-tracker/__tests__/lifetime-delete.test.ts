import assert from 'node:assert/strict';
import test from 'node:test';
import { startSessionRecord } from '../session.js';
import { applySessionLifetimeSummary, rebuildLifetimeSummaries } from '../lifetime.js';
import { deleteMaintenanceBatch } from '../query-delete-maintenance.js';
import { toDbTimestamp } from '../query-shared.js';
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

test('fractional lifetime metrics stay normalized across apply, rebuild, and delete', () => {
  const db = createDb();
  try {
    const videoId = seedVideo(db, null, 'fractional-metrics');
    const seedFractionalSession = (
      startedAtMs: number,
      metrics: { activeMs: number; cards: number; lines: number; tokens: number },
    ) => {
      const { state } = startSessionRecord(db, videoId, startedAtMs);
      state.activeWatchedMs = metrics.activeMs;
      state.cardsMined = metrics.cards;
      state.linesSeen = metrics.lines;
      state.tokensSeen = metrics.tokens;
      const endedAtMs = startedAtMs + 2_000;
      db.prepare(
        `UPDATE imm_sessions SET
           ended_at_ms = ?,
           active_watched_ms = ?,
           cards_mined = ?,
           lines_seen = ?,
           tokens_seen = ?
         WHERE session_id = ?`,
      ).run(
        toDbTimestamp(endedAtMs),
        metrics.activeMs,
        metrics.cards,
        metrics.lines,
        metrics.tokens,
        state.sessionId,
      );
      return { state, endedAtMs };
    };
    const readMediaMetrics = () =>
      cleanRow<{
        total_sessions: number;
        total_active_ms: number;
        total_cards: number;
        total_lines_seen: number;
        total_tokens_seen: number;
      }>(
        db
          .prepare(
            `SELECT total_sessions, total_active_ms, total_cards,
                    total_lines_seen, total_tokens_seen
             FROM imm_lifetime_media WHERE video_id = ?`,
          )
          .get(videoId),
      );

    const withoutTelemetry = seedFractionalSession(BASE_MS, {
      activeMs: 1_234.9,
      cards: 2.8,
      lines: 3.7,
      tokens: 4.6,
    });
    applySessionLifetimeSummary(db, withoutTelemetry.state, withoutTelemetry.endedAtMs);
    assert.deepEqual(readMediaMetrics(), {
      total_sessions: 1,
      total_active_ms: 1_234,
      total_cards: 2,
      total_lines_seen: 3,
      total_tokens_seen: 4,
    });

    const withTelemetry = seedFractionalSession(BASE_MS + DAY_MS, {
      activeMs: 9_999.9,
      cards: 9.9,
      lines: 9.9,
      tokens: 9.9,
    });
    db.prepare(
      `INSERT INTO imm_session_telemetry (
         session_id, sample_ms, active_watched_ms, cards_mined, lines_seen, tokens_seen
       ) VALUES (?, ?, ?, ?, ?, ?)`,
    ).run(withTelemetry.state.sessionId, withTelemetry.endedAtMs, 2_345.9, 5.8, 6.7, 7.6);
    applySessionLifetimeSummary(db, withTelemetry.state, withTelemetry.endedAtMs);
    assert.deepEqual(readMediaMetrics(), {
      total_sessions: 2,
      total_active_ms: 3_579,
      total_cards: 7,
      total_lines_seen: 9,
      total_tokens_seen: 11,
    });

    deleteMaintenanceBatch(db, [{ kind: 'session', sessionId: withTelemetry.state.sessionId }]);
    const retainedMetrics = {
      total_sessions: 1,
      total_active_ms: 1_234,
      total_cards: 2,
      total_lines_seen: 3,
      total_tokens_seen: 4,
    };
    assert.deepEqual(readMediaMetrics(), retainedMetrics, 'delete subtracts floored telemetry');

    rebuildLifetimeSummaries(db);
    assert.deepEqual(
      readMediaMetrics(),
      retainedMetrics,
      'rebuild floors session-row fallback values',
    );

    deleteMaintenanceBatch(db, [{ kind: 'session', sessionId: withoutTelemetry.state.sessionId }]);
    assert.deepEqual(snapshotMedia(db), [], 'delete subtracts the normalized metrics exactly');
    assert.deepEqual(snapshotGlobal(db), {
      total_sessions: 0,
      total_active_ms: 0,
      total_cards: 0,
      active_days: 0,
      episodes_started: 0,
      episodes_completed: 0,
      anime_completed: 0,
    });
  } finally {
    db.close();
  }
});

test('incremental delete maintenance matches a full rebuild when no history is pruned', () => {
  const db = createDb();
  try {
    const animeA = seedAnime(db, 'Anime A', 2);
    const animeB = seedAnime(db, 'Anime B', 1);
    const videoA1 = seedVideo(db, animeA, 'anime-a-ep1', { watched: true });
    const videoA2 = seedVideo(db, animeA, 'anime-a-ep2', { watched: true });
    const videoB1 = seedVideo(db, animeB, 'anime-b-ep1', { watched: true });
    const videoLoose = seedVideo(db, null, 'loose-video');

    seedEndedSession(db, videoA1, BASE_MS, { activeMs: 60_000, cards: 2, lines: 30, tokens: 200 });
    const deletedSessionId = seedEndedSession(db, videoA1, BASE_MS + DAY_MS, {
      activeMs: 45_000,
      cards: 1,
      lines: 20,
      tokens: 100,
    });
    seedEndedSession(db, videoA2, BASE_MS + 2 * DAY_MS, { activeMs: 90_000, cards: 3 });
    seedEndedSession(db, videoB1, BASE_MS + 3 * DAY_MS, { activeMs: 30_000 });
    seedEndedSession(db, videoLoose, BASE_MS + 4 * DAY_MS, { activeMs: 15_000 });

    rebuildLifetimeSummaries(db);

    deleteMaintenanceBatch(db, [
      { kind: 'session', sessionId: deletedSessionId },
      { kind: 'video', videoId: videoLoose },
      { kind: 'anime', animeId: animeB },
    ]);

    const incrementalGlobal = snapshotGlobal(db);
    const incrementalMedia = snapshotMedia(db);
    const incrementalAnime = snapshotAnime(db);

    // With every session still retained, subtracting must land on exactly the
    // state a from-scratch rebuild computes.
    rebuildLifetimeSummaries(db);
    assert.deepEqual(incrementalGlobal, snapshotGlobal(db));
    assert.deepEqual(incrementalMedia, snapshotMedia(db));
    assert.deepEqual(incrementalAnime, snapshotAnime(db));
  } finally {
    db.close();
  }
});

test('deleting a retained session preserves lifetime history from pruned sessions', () => {
  const db = createDb();
  try {
    const animeId = seedAnime(db, 'Pruned Anime', null);
    const videoId = seedVideo(db, animeId, 'pruned-ep1');
    const prunedSessionId = seedEndedSession(db, videoId, BASE_MS, {
      activeMs: 120_000,
      cards: 4,
      lines: 50,
      tokens: 400,
    });
    const retainedSessionId = seedEndedSession(db, videoId, BASE_MS + DAY_MS, {
      activeMs: 30_000,
      cards: 1,
      lines: 10,
      tokens: 80,
    });
    rebuildLifetimeSummaries(db);

    // Simulate raw-session retention pruning the older session. Lifetime
    // summaries intentionally keep its contribution.
    db.prepare('DELETE FROM imm_sessions WHERE session_id = ?').run(prunedSessionId);

    deleteMaintenanceBatch(db, [{ kind: 'session', sessionId: retainedSessionId }]);

    const globalRow = snapshotGlobal(db);
    assert.equal(globalRow.total_sessions, 1, 'pruned session contribution survives the delete');
    assert.equal(globalRow.total_active_ms, 120_000);
    assert.equal(globalRow.total_cards, 4);
    assert.equal(globalRow.episodes_started, 1);
    // The pruned session's day stays counted (pruning never subtracts); only
    // the deleted retained session's day is dropped.
    assert.equal(globalRow.active_days, 1);

    const mediaRow = db
      .prepare(
        'SELECT total_sessions, total_active_ms, total_cards FROM imm_lifetime_media WHERE video_id = ?',
      )
      .get(videoId);
    assert.deepEqual(
      cleanRow<{ total_sessions: number; total_active_ms: number; total_cards: number }>(mediaRow),
      { total_sessions: 1, total_active_ms: 120_000, total_cards: 4 },
    );
  } finally {
    db.close();
  }
});

test('active_days only drops when the last session of a local day is deleted', () => {
  const db = createDb();
  try {
    const videoId = seedVideo(db, null, 'same-day');
    const firstSessionId = seedEndedSession(db, videoId, BASE_MS, { activeMs: 10_000 });
    const secondSessionId = seedEndedSession(db, videoId, BASE_MS + 3_600_000, {
      activeMs: 20_000,
    });
    rebuildLifetimeSummaries(db);
    assert.equal(snapshotGlobal(db).active_days, 1);

    deleteMaintenanceBatch(db, [{ kind: 'session', sessionId: firstSessionId }]);
    assert.equal(snapshotGlobal(db).active_days, 1, 'day still has a session');

    deleteMaintenanceBatch(db, [{ kind: 'session', sessionId: secondSessionId }]);
    assert.equal(snapshotGlobal(db).active_days, 0, 'day lost its last session');
  } finally {
    db.close();
  }
});

test('deleting a video updates anime and global rollups without a rebuild', () => {
  const db = createDb();
  try {
    const animeId = seedAnime(db, 'Two Episode Anime', 2);
    const videoEp1 = seedVideo(db, animeId, 'two-ep-1', { watched: true });
    const videoEp2 = seedVideo(db, animeId, 'two-ep-2', { watched: true });
    seedEndedSession(db, videoEp1, BASE_MS, { activeMs: 60_000, cards: 2 });
    seedEndedSession(db, videoEp2, BASE_MS + DAY_MS, { activeMs: 40_000, cards: 1 });
    rebuildLifetimeSummaries(db);
    assert.equal(snapshotGlobal(db).anime_completed, 1);

    deleteMaintenanceBatch(db, [{ kind: 'video', videoId: videoEp2 }]);

    const globalRow = snapshotGlobal(db);
    assert.equal(globalRow.total_sessions, 1);
    assert.equal(globalRow.total_active_ms, 60_000);
    assert.equal(globalRow.episodes_started, 1);
    assert.equal(globalRow.episodes_completed, 1);
    assert.equal(globalRow.anime_completed, 0, 'anime no longer has all episodes completed');

    const animeRow = db
      .prepare(
        'SELECT total_sessions, episodes_started, episodes_completed FROM imm_lifetime_anime WHERE anime_id = ?',
      )
      .get(animeId);
    assert.deepEqual(
      cleanRow<{
        total_sessions: number;
        episodes_started: number;
        episodes_completed: number;
      }>(animeRow),
      { total_sessions: 1, episodes_started: 1, episodes_completed: 1 },
    );

    deleteMaintenanceBatch(db, [{ kind: 'video', videoId: videoEp1 }]);
    assert.equal(
      (db.prepare('SELECT COUNT(*) AS total FROM imm_lifetime_anime').get() as { total: number })
        .total,
      0,
      'anime lifetime row is dropped once no episodes remain',
    );
    assert.deepEqual(snapshotGlobal(db), {
      total_sessions: 0,
      total_active_ms: 0,
      total_cards: 0,
      active_days: 0,
      episodes_started: 0,
      episodes_completed: 0,
      anime_completed: 0,
    });
  } finally {
    db.close();
  }
});
