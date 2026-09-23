import type { DatabaseSync } from './sqlite';
import { fromDbTimestamp } from './query-shared';

/** Watch state of one video row, addressed by the key playback recorded it under. */
export interface VideoWatchStateRow {
  videoKey: string;
  watched: boolean;
  /** Start of the most recent session on this video, or null when never played. */
  lastWatchedMs: number | null;
  sessionCount: number;
}

/** The video row a key belongs to, or null when nothing has recorded it yet. */
export function getVideoIdByVideoKey(db: DatabaseSync, videoKey: string): number | null {
  const row = db.prepare('SELECT video_id FROM imm_videos WHERE video_key = ?').get(videoKey) as {
    video_id: number;
  } | null;
  return row?.video_id ?? null;
}

/**
 * How many keys one statement binds. SQLite caps parameters per statement, and
 * an episode list can be long, so the lookup runs in chunks.
 */
const CHUNK_SIZE = 400;

/**
 * Look up watch state for a set of video keys.
 *
 * Keys that have never been played simply do not come back — the caller treats
 * a missing key as unwatched rather than needing a row for it.
 *
 * `last_watched_ms` on the lifetime tables is rebuilt in batches and can lag, so
 * the timestamp comes from the sessions themselves. Timestamps are epoch
 * milliseconds stored as text, hence the cast before `MAX`.
 */
export function getWatchStateByVideoKeys(
  db: DatabaseSync,
  videoKeys: string[],
): VideoWatchStateRow[] {
  const unique = [...new Set(videoKeys.filter((key) => key.length > 0))];
  const rows: VideoWatchStateRow[] = [];

  for (let start = 0; start < unique.length; start += CHUNK_SIZE) {
    const chunk = unique.slice(start, start + CHUNK_SIZE);
    const placeholders = chunk.map(() => '?').join(', ');
    const chunkRows = db
      .prepare(
        `
    SELECT
      v.video_key AS videoKey,
      v.watched AS watched,
      MAX(CAST(s.started_at_ms AS INTEGER)) AS lastWatchedMs,
      COUNT(s.session_id) AS sessionCount
    FROM imm_videos v
    LEFT JOIN imm_sessions s ON s.video_id = v.video_id
    WHERE v.video_key IN (${placeholders})
    GROUP BY v.video_id
  `,
      )
      .all(...chunk) as Array<{
      videoKey: string;
      watched: number;
      lastWatchedMs: number | string | null;
      sessionCount: number;
    }>;

    for (const row of chunkRows) {
      rows.push({
        videoKey: row.videoKey,
        watched: row.watched === 1,
        lastWatchedMs: fromDbTimestamp(row.lastWatchedMs),
        sessionCount: Number(row.sessionCount) || 0,
      });
    }
  }

  return rows;
}
