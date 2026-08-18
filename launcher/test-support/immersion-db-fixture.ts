import { Database } from 'bun:sqlite';
import { SCHEMA_VERSION } from '../../src/core/services/immersion-tracker/types.js';
import { IMMERSION_DB_FIXTURE_DDL } from './immersion-db-schema.js';

export function createImmersionDbFixture(dbPath: string): void {
  const db = new Database(dbPath, { create: true });
  try {
    db.run('PRAGMA foreign_keys = ON');
    db.run('PRAGMA journal_mode = WAL');
    for (const statement of IMMERSION_DB_FIXTURE_DDL.split(';')) {
      const sql = statement.trim();
      if (sql) db.run(sql);
    }
    db.prepare('INSERT INTO imm_schema_version (schema_version, applied_at_ms) VALUES (?, ?)').run(
      SCHEMA_VERSION,
      String(Date.now()),
    );
    db.prepare(
      `INSERT INTO imm_rollup_state(state_key, state_value) VALUES ('last_rollup_sample_ms', 0)`,
    ).run();
    db.prepare(
      `INSERT INTO imm_rollup_state(state_key, state_value) VALUES ('lexical_daily_rollups_version', 0)`,
    ).run();
    db.prepare(
      `INSERT INTO imm_lifetime_global(global_id, CREATED_DATE, LAST_UPDATE_DATE) VALUES (1, ?, ?)`,
    ).run(String(Date.now()), String(Date.now()));
  } finally {
    db.close();
  }
}

let uniqueCounter = 0;

export interface FixtureSessionInput {
  uuid: string;
  videoKey: string;
  animeTitleKey?: string | null;
  animeEpisodesTotal?: number | null;
  startedAtMs: number;
  endedAtMs?: number | null;
  activeWatchedMs?: number;
  cardsMined?: number;
  linesSeen?: number;
  tokensSeen?: number;
  watched?: boolean;
  words?: Array<{ headword: string; word: string; reading: string; count: number }>;
  applyLifetime?: boolean;
}

/**
 * Insert a session (with video/anime/subtitle-line/occurrence rows) the way
 * the app would have recorded it, optionally crediting lifetime aggregates
 * the way applySessionLifetimeSummary does for locally-recorded sessions.
 */
export function insertFixtureSession(dbPath: string, input: FixtureSessionInput): void {
  const db = new Database(dbPath, { readwrite: true });
  const stamp = String(input.startedAtMs);
  try {
    db.run('PRAGMA foreign_keys = ON');
    let animeId: number | null = null;
    if (input.animeTitleKey) {
      const existing = db
        .query<{
          anime_id: number;
        }>('SELECT anime_id FROM imm_anime WHERE normalized_title_key = ?')
        .get(input.animeTitleKey);
      animeId = existing
        ? existing.anime_id
        : Number(
            db
              .prepare(
                `INSERT INTO imm_anime (normalized_title_key, canonical_title, episodes_total, CREATED_DATE, LAST_UPDATE_DATE)
                 VALUES (?, ?, ?, ?, ?)`,
              )
              .run(
                input.animeTitleKey,
                input.animeTitleKey,
                input.animeEpisodesTotal ?? null,
                stamp,
                stamp,
              ).lastInsertRowid,
          );
    }

    const existingVideo = db
      .query<{ video_id: number }>('SELECT video_id FROM imm_videos WHERE video_key = ?')
      .get(input.videoKey);
    const videoId = existingVideo
      ? existingVideo.video_id
      : Number(
          db
            .prepare(
              `INSERT INTO imm_videos (video_key, anime_id, canonical_title, source_type, source_path, watched, duration_ms, CREATED_DATE, LAST_UPDATE_DATE)
               VALUES (?, ?, ?, 1, ?, ?, 1440000, ?, ?)`,
            )
            .run(
              input.videoKey,
              animeId,
              input.videoKey,
              `/videos/${input.videoKey}.mkv`,
              input.watched ? 1 : 0,
              stamp,
              stamp,
            ).lastInsertRowid,
        );
    if (input.watched) {
      db.prepare('UPDATE imm_videos SET watched = 1 WHERE video_id = ?').run(videoId);
    }

    const endedAtMs =
      input.endedAtMs === undefined ? input.startedAtMs + 1_500_000 : input.endedAtMs;
    const activeWatchedMs = input.activeWatchedMs ?? 1_200_000;
    const cardsMined = input.cardsMined ?? 2;
    const linesSeen = input.linesSeen ?? 100;
    const tokensSeen = input.tokensSeen ?? 800;
    const sessionId = Number(
      db
        .prepare(
          `INSERT INTO imm_sessions (
             session_uuid, video_id, started_at_ms, ended_at_ms, status,
             total_watched_ms, active_watched_ms, lines_seen, tokens_seen, cards_mined,
             CREATED_DATE, LAST_UPDATE_DATE
           ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        )
        .run(
          input.uuid,
          videoId,
          String(input.startedAtMs),
          endedAtMs === null ? null : String(endedAtMs),
          endedAtMs === null ? 1 : 2,
          activeWatchedMs,
          activeWatchedMs,
          linesSeen,
          tokensSeen,
          cardsMined,
          stamp,
          stamp,
        ).lastInsertRowid,
    );

    db.prepare(
      `INSERT INTO imm_session_telemetry (
         session_id, sample_ms, total_watched_ms, active_watched_ms, lines_seen, tokens_seen,
         cards_mined, CREATED_DATE, LAST_UPDATE_DATE
       ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    ).run(
      sessionId,
      String(endedAtMs ?? input.startedAtMs),
      activeWatchedMs,
      activeWatchedMs,
      linesSeen,
      tokensSeen,
      cardsMined,
      stamp,
      stamp,
    );

    for (const word of input.words ?? []) {
      const existing = db
        .query<{
          id: number;
        }>('SELECT id FROM imm_words WHERE headword = ? AND word = ? AND reading = ?')
        .get(word.headword, word.word, word.reading);
      const wordId = existing
        ? existing.id
        : Number(
            db
              .prepare(
                `INSERT INTO imm_words (headword, word, reading, first_seen, last_seen, frequency)
                 VALUES (?, ?, ?, ?, ?, 0)`,
              )
              .run(
                word.headword,
                word.word,
                word.reading,
                Math.floor(input.startedAtMs / 1000),
                Math.floor(input.startedAtMs / 1000),
              ).lastInsertRowid,
          );
      uniqueCounter += 1;
      const lineId = Number(
        db
          .prepare(
            `INSERT INTO imm_subtitle_lines (session_id, video_id, anime_id, line_index, text, CREATED_DATE, LAST_UPDATE_DATE)
             VALUES (?, ?, ?, ?, ?, ?, ?)`,
          )
          .run(
            sessionId,
            videoId,
            animeId,
            uniqueCounter,
            `line ${word.word}`,
            input.startedAtMs,
            input.startedAtMs,
          ).lastInsertRowid,
      );
      db.prepare(
        'INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count) VALUES (?, ?, ?)',
      ).run(lineId, wordId, word.count);
      db.prepare('UPDATE imm_words SET frequency = frequency + ? WHERE id = ?').run(
        word.count,
        wordId,
      );
    }

    if (input.applyLifetime && endedAtMs !== null) {
      applyFixtureLifetime(db, sessionId, videoId, animeId, {
        endedAtMs,
        startedAtMs: input.startedAtMs,
        activeWatchedMs,
        cardsMined,
        linesSeen,
        tokensSeen,
        watched: Boolean(input.watched),
        episodesTotal: input.animeEpisodesTotal ?? null,
      });
    }
  } finally {
    db.close();
  }
}

function applyFixtureLifetime(
  db: Database,
  sessionId: number,
  videoId: number,
  animeId: number | null,
  data: {
    endedAtMs: number;
    startedAtMs: number;
    activeWatchedMs: number;
    cardsMined: number;
    linesSeen: number;
    tokensSeen: number;
    watched: boolean;
    episodesTotal: number | null;
  },
): void {
  const stamp = String(data.endedAtMs);
  db.prepare(
    `INSERT INTO imm_lifetime_applied_sessions (session_id, applied_at_ms, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, ?, ?)`,
  ).run(sessionId, String(data.endedAtMs), stamp, stamp);

  const mediaLifetime = db
    .query<{ completed: number }>('SELECT completed FROM imm_lifetime_media WHERE video_id = ?')
    .get(videoId);
  const isFirstSessionForVideo = !mediaLifetime;
  const isFirstCompleted = data.watched && Number(mediaLifetime?.completed ?? 0) <= 0;
  const dayExpr = `CAST(julianday(CAST(started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER)`;
  const otherOnDay = db
    .query(
      `SELECT 1 FROM imm_sessions WHERE session_id != ? AND ${dayExpr} =
         CAST(julianday(CAST(? AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) LIMIT 1`,
    )
    .get(sessionId, String(data.startedAtMs));

  db.prepare(
    `UPDATE imm_lifetime_global SET
       total_sessions = total_sessions + 1,
       total_active_ms = total_active_ms + ?,
       total_cards = total_cards + ?,
       active_days = active_days + ?,
       episodes_started = episodes_started + ?,
       episodes_completed = episodes_completed + ?
     WHERE global_id = 1`,
  ).run(
    data.activeWatchedMs,
    data.cardsMined,
    otherOnDay ? 0 : 1,
    isFirstSessionForVideo ? 1 : 0,
    isFirstCompleted ? 1 : 0,
  );

  db.prepare(
    `INSERT INTO imm_lifetime_media (video_id, total_sessions, total_active_ms, total_cards, total_lines_seen, total_tokens_seen, completed, first_watched_ms, last_watched_ms, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?)
     ON CONFLICT(video_id) DO UPDATE SET
       total_sessions = total_sessions + 1,
       total_active_ms = total_active_ms + excluded.total_active_ms,
       total_cards = total_cards + excluded.total_cards,
       total_lines_seen = total_lines_seen + excluded.total_lines_seen,
       total_tokens_seen = total_tokens_seen + excluded.total_tokens_seen,
       completed = MAX(completed, excluded.completed),
       last_watched_ms = excluded.last_watched_ms`,
  ).run(
    videoId,
    data.activeWatchedMs,
    data.cardsMined,
    data.linesSeen,
    data.tokensSeen,
    data.watched ? 1 : 0,
    String(data.startedAtMs),
    String(data.endedAtMs),
    stamp,
    stamp,
  );

  if (animeId !== null) {
    db.prepare(
      `INSERT INTO imm_lifetime_anime (anime_id, total_sessions, total_active_ms, total_cards, total_lines_seen, total_tokens_seen, episodes_started, episodes_completed, first_watched_ms, last_watched_ms, CREATED_DATE, LAST_UPDATE_DATE)
       VALUES (?, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
       ON CONFLICT(anime_id) DO UPDATE SET
         total_sessions = total_sessions + 1,
         total_active_ms = total_active_ms + excluded.total_active_ms,
         total_cards = total_cards + excluded.total_cards,
         total_lines_seen = total_lines_seen + excluded.total_lines_seen,
         total_tokens_seen = total_tokens_seen + excluded.total_tokens_seen,
         episodes_started = episodes_started + excluded.episodes_started,
         episodes_completed = episodes_completed + excluded.episodes_completed,
         last_watched_ms = excluded.last_watched_ms`,
    ).run(
      animeId,
      data.activeWatchedMs,
      data.cardsMined,
      data.linesSeen,
      data.tokensSeen,
      isFirstSessionForVideo ? 1 : 0,
      isFirstCompleted ? 1 : 0,
      String(data.startedAtMs),
      String(data.endedAtMs),
      stamp,
      stamp,
    );
  }
}
