import crypto from 'node:crypto';
import type { DatabaseSync } from './sqlite';
import { createInitialSessionState } from './reducer';
import { SESSION_STATUS_ACTIVE, SESSION_STATUS_ENDED } from './types';
import type { SessionState } from './types';

export function startSessionRecord(
  db: DatabaseSync,
  videoId: number,
  startedAtMs = Date.now(),
): { sessionId: number; state: SessionState } {
  const sessionUuid = crypto.randomUUID();
  const nowMs = Date.now();
  const result = db
    .prepare(
      `
        INSERT INTO imm_sessions (
          session_uuid, video_id, started_at_ms, status,
          CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (?, ?, ?, ?, ?, ?)
    `,
    )
    .run(sessionUuid, videoId, startedAtMs, SESSION_STATUS_ACTIVE, startedAtMs, nowMs);
  const sessionId = Number(result.lastInsertRowid);
  return {
    sessionId,
    state: createInitialSessionState(sessionId, videoId, startedAtMs),
  };
}

export function finalizeSessionRecord(
  db: DatabaseSync,
  sessionState: SessionState,
  endedAtMs = Date.now(),
): void {
  db.prepare(
    `
      UPDATE imm_sessions
      SET
        ended_at_ms = ?,
        status = ?,
        ended_media_ms = ?,
        total_watched_ms = ?,
        active_watched_ms = ?,
        lines_seen = ?,
        tokens_seen = ?,
        cards_mined = ?,
        lookup_count = ?,
        lookup_hits = ?,
        yomitan_lookup_count = ?,
        pause_count = ?,
        pause_ms = ?,
        seek_forward_count = ?,
        seek_backward_count = ?,
        media_buffer_events = ?,
        LAST_UPDATE_DATE = ?
      WHERE session_id = ?
    `,
  ).run(
    endedAtMs,
    SESSION_STATUS_ENDED,
    sessionState.lastMediaMs,
    sessionState.totalWatchedMs,
    sessionState.activeWatchedMs,
    sessionState.linesSeen,
    sessionState.tokensSeen,
    sessionState.cardsMined,
    sessionState.lookupCount,
    sessionState.lookupHits,
    sessionState.yomitanLookupCount,
    sessionState.pauseCount,
    sessionState.pauseMs,
    sessionState.seekForwardCount,
    sessionState.seekBackwardCount,
    sessionState.mediaBufferEvents,
    Date.now(),
    sessionState.sessionId,
  );
}
