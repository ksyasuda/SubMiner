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
        LAST_UPDATE_DATE = ?
      WHERE session_id = ?
    `,
  ).run(endedAtMs, SESSION_STATUS_ENDED, Date.now(), sessionState.sessionId);
}
