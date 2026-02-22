import crypto from 'node:crypto';
import type { DatabaseSync } from 'node:sqlite';
import { createInitialSessionState } from './reducer';
import { SESSION_STATUS_ACTIVE, SESSION_STATUS_ENDED } from './types';
import type { SessionState } from './types';

export function startSessionRecord(
  db: DatabaseSync,
  videoId: number,
  startedAtMs = Date.now(),
): { sessionId: number; state: SessionState } {
  const sessionUuid = crypto.randomUUID();
  const result = db
    .prepare(
      `
        INSERT INTO imm_sessions (
          session_uuid, video_id, started_at_ms, status, created_at_ms, updated_at_ms
        ) VALUES (?, ?, ?, ?, ?, ?)
      `,
    )
    .run(sessionUuid, videoId, startedAtMs, SESSION_STATUS_ACTIVE, startedAtMs, startedAtMs);
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
    'UPDATE imm_sessions SET ended_at_ms = ?, status = ?, updated_at_ms = ? WHERE session_id = ?',
  ).run(endedAtMs, SESSION_STATUS_ENDED, Date.now(), sessionState.sessionId);
}
