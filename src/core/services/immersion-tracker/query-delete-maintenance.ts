import type { DatabaseSync } from './sqlite';
import { rebuildLifetimeSummariesInTransaction } from './lifetime';
import { getRollupGroupsForSessions, refreshRollupsForGroupsInTransaction } from './maintenance';
import {
  applyLexicalRemovals,
  cleanupUnusedCoverArtBlobHash,
  deleteSessionsByIds,
  makePlaceholders,
  planLexicalRemovalsForSessions,
} from './query-shared';

export type DeleteMaintenanceOperation =
  | { kind: 'session'; sessionId: number }
  | { kind: 'sessions'; sessionIds: number[] }
  | { kind: 'video'; videoId: number }
  | { kind: 'anime'; animeId: number };

function addOperationTargets(
  operations: DeleteMaintenanceOperation[],
  sessionIds: Set<number>,
  videoIds: Set<number>,
  animeIds: Set<number>,
): void {
  for (const operation of operations) {
    switch (operation.kind) {
      case 'session':
        sessionIds.add(operation.sessionId);
        break;
      case 'sessions':
        for (const sessionId of operation.sessionIds) sessionIds.add(sessionId);
        break;
      case 'video':
        videoIds.add(operation.videoId);
        break;
      case 'anime':
        animeIds.add(operation.animeId);
        break;
    }
  }
}

function selectIds(db: DatabaseSync, sql: string, params: number[], column: string): number[] {
  if (params.length === 0) return [];
  return (db.prepare(sql).all(...params) as Array<Record<string, number>>).map(
    (row) => row[column]!,
  );
}

export function deleteMaintenanceBatch(
  db: DatabaseSync,
  operations: DeleteMaintenanceOperation[],
): void {
  if (operations.length === 0) return;

  db.exec('BEGIN IMMEDIATE');
  try {
    const sessionIds = new Set<number>();
    const videoIds = new Set<number>();
    const animeIds = new Set<number>();
    addOperationTargets(operations, sessionIds, videoIds, animeIds);

    const animeIdList = [...animeIds];
    for (const videoId of selectIds(
      db,
      `SELECT video_id FROM imm_videos WHERE anime_id IN (${makePlaceholders(animeIdList)})`,
      animeIdList,
      'video_id',
    )) {
      videoIds.add(videoId);
    }

    const videoIdList = [...videoIds];
    for (const sessionId of selectIds(
      db,
      `SELECT session_id FROM imm_sessions WHERE video_id IN (${makePlaceholders(videoIdList)})`,
      videoIdList,
      'session_id',
    )) {
      sessionIds.add(sessionId);
    }

    const sessionIdList = [...sessionIds];
    const lexicalRemovals = planLexicalRemovalsForSessions(db, sessionIdList);
    const affectedRollupGroups = getRollupGroupsForSessions(db, sessionIdList).filter(
      (group) => !videoIds.has(group.videoId),
    );
    const coverBlobHashes = new Set<string>();
    if (videoIdList.length > 0) {
      const placeholders = makePlaceholders(videoIdList);
      const artRows = db
        .prepare(
          `SELECT cover_blob_hash AS coverBlobHash
           FROM imm_media_art
           WHERE video_id IN (${placeholders}) AND cover_blob_hash IS NOT NULL`,
        )
        .all(...videoIdList) as Array<{ coverBlobHash: string }>;
      for (const row of artRows) coverBlobHashes.add(row.coverBlobHash);

      deleteSessionsByIds(db, sessionIdList);
      db.prepare(`DELETE FROM imm_subtitle_lines WHERE video_id IN (${placeholders})`).run(
        ...videoIdList,
      );
      db.prepare(`DELETE FROM imm_daily_rollups WHERE video_id IN (${placeholders})`).run(
        ...videoIdList,
      );
      db.prepare(`DELETE FROM imm_monthly_rollups WHERE video_id IN (${placeholders})`).run(
        ...videoIdList,
      );
      db.prepare(`DELETE FROM imm_media_art WHERE video_id IN (${placeholders})`).run(
        ...videoIdList,
      );
      db.prepare(`DELETE FROM imm_videos WHERE video_id IN (${placeholders})`).run(...videoIdList);
    } else {
      deleteSessionsByIds(db, sessionIdList);
    }

    for (const coverBlobHash of coverBlobHashes) {
      cleanupUnusedCoverArtBlobHash(db, coverBlobHash);
    }
    if (animeIdList.length > 0) {
      const placeholders = makePlaceholders(animeIdList);
      db.prepare(`DELETE FROM imm_lifetime_anime WHERE anime_id IN (${placeholders})`).run(
        ...animeIdList,
      );
      db.prepare(`DELETE FROM imm_anime WHERE anime_id IN (${placeholders})`).run(...animeIdList);
    }

    applyLexicalRemovals(db, lexicalRemovals);
    rebuildLifetimeSummariesInTransaction(db);
    refreshRollupsForGroupsInTransaction(db, affectedRollupGroups);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}
