import type { DatabaseSync } from './sqlite';
import { rebuildLifetimeSummariesInTransaction } from './lifetime';
import { getRollupGroupsForSessions, refreshRollupsForGroupsInTransaction } from './maintenance';
import {
  applyLexicalRemovals,
  cleanupUnusedCoverArtBlobHash,
  deleteSessionsByIds,
  forEachIdChunk,
  makePlaceholders,
  planLexicalRemovalsForSessions,
  SQLITE_ID_CHUNK_SIZE,
  type LexicalRemovalPlan,
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

function selectIds(
  db: DatabaseSync,
  buildSql: (placeholders: string) => string,
  params: number[],
  column: string,
): number[] {
  if (params.length === 0) return [];
  const ids: number[] = [];
  forEachIdChunk(params, (chunk) => {
    const rows = db.prepare(buildSql(makePlaceholders(chunk))).all(...chunk) as Array<
      Record<string, number>
    >;
    for (const row of rows) ids.push(row[column]!);
  });
  return ids;
}

function planLexicalRemovalsInChunks(db: DatabaseSync, sessionIds: number[]): LexicalRemovalPlan {
  const combined: LexicalRemovalPlan = { words: [], kanji: [] };
  const merge = (target: LexicalRemovalPlan['words'], source: LexicalRemovalPlan['words']) => {
    const byId = new Map(target.map((entry) => [entry.id, entry]));
    for (const entry of source) {
      const existing = byId.get(entry.id);
      if (!existing) {
        const added = { ...entry };
        target.push(added);
        byId.set(entry.id, added);
        continue;
      }
      existing.removedFrequency += entry.removedFrequency;
      if (
        entry.removedFirstSeenMs !== null &&
        (existing.removedFirstSeenMs === null ||
          entry.removedFirstSeenMs < existing.removedFirstSeenMs)
      ) {
        existing.removedFirstSeenMs = entry.removedFirstSeenMs;
      }
      if (
        entry.removedLastSeenMs !== null &&
        (existing.removedLastSeenMs === null ||
          entry.removedLastSeenMs > existing.removedLastSeenMs)
      ) {
        existing.removedLastSeenMs = entry.removedLastSeenMs;
      }
    }
  };

  forEachIdChunk(sessionIds, (chunk) => {
    const plan = planLexicalRemovalsForSessions(db, chunk);
    merge(combined.words, plan.words);
    merge(combined.kanji, plan.kanji);
  });
  return combined;
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
      (placeholders) => `SELECT video_id FROM imm_videos WHERE anime_id IN (${placeholders})`,
      animeIdList,
      'video_id',
    )) {
      videoIds.add(videoId);
    }

    const videoIdList = [...videoIds];
    for (const sessionId of selectIds(
      db,
      (placeholders) => `SELECT session_id FROM imm_sessions WHERE video_id IN (${placeholders})`,
      videoIdList,
      'session_id',
    )) {
      sessionIds.add(sessionId);
    }

    const sessionIdList = [...sessionIds];
    const lexicalRemovals = planLexicalRemovalsInChunks(db, sessionIdList);
    const affectedRollupGroups = sessionIdList
      .flatMap((_, index) =>
        index % SQLITE_ID_CHUNK_SIZE === 0
          ? getRollupGroupsForSessions(db, sessionIdList.slice(index, index + SQLITE_ID_CHUNK_SIZE))
          : [],
      )
      .filter((group) => !videoIds.has(group.videoId));
    const coverBlobHashes = new Set<string>();
    if (videoIdList.length > 0) {
      forEachIdChunk(videoIdList, (chunk) => {
        const placeholders = makePlaceholders(chunk);
        const artRows = db
          .prepare(
            `SELECT cover_blob_hash AS coverBlobHash
             FROM imm_media_art
             WHERE video_id IN (${placeholders}) AND cover_blob_hash IS NOT NULL`,
          )
          .all(...chunk) as Array<{ coverBlobHash: string }>;
        for (const row of artRows) coverBlobHashes.add(row.coverBlobHash);
      });

      deleteSessionsByIds(db, sessionIdList);
      forEachIdChunk(videoIdList, (chunk) => {
        const placeholders = makePlaceholders(chunk);
        db.prepare(`DELETE FROM imm_subtitle_lines WHERE video_id IN (${placeholders})`).run(
          ...chunk,
        );
        db.prepare(`DELETE FROM imm_daily_rollups WHERE video_id IN (${placeholders})`).run(
          ...chunk,
        );
        db.prepare(`DELETE FROM imm_monthly_rollups WHERE video_id IN (${placeholders})`).run(
          ...chunk,
        );
        db.prepare(`DELETE FROM imm_media_art WHERE video_id IN (${placeholders})`).run(...chunk);
        db.prepare(`DELETE FROM imm_videos WHERE video_id IN (${placeholders})`).run(...chunk);
      });
    } else {
      deleteSessionsByIds(db, sessionIdList);
    }

    for (const coverBlobHash of coverBlobHashes) {
      cleanupUnusedCoverArtBlobHash(db, coverBlobHash);
    }
    if (animeIdList.length > 0) {
      forEachIdChunk(animeIdList, (chunk) => {
        const placeholders = makePlaceholders(chunk);
        db.prepare(`DELETE FROM imm_lifetime_anime WHERE anime_id IN (${placeholders})`).run(
          ...chunk,
        );
        db.prepare(`DELETE FROM imm_anime WHERE anime_id IN (${placeholders})`).run(...chunk);
      });
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
