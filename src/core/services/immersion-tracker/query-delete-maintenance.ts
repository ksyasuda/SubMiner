import type { DatabaseSync } from './sqlite';
import { applyLifetimeRemovals, planLifetimeRemovals } from './lifetime';
import { getRollupGroupsForSessions, refreshRollupsForGroupsInTransaction } from './maintenance';
import {
  applyLexicalRemovals,
  cleanupUnusedCoverArtBlobHash,
  deleteSessionsByIds,
  forEachIdChunk,
  makePlaceholders,
  planLexicalRemovalsForSessions,
  planLexicalRemovalsForVideos,
  type LexicalRemovalPlan,
} from './query-shared';
import type { RollupGroup } from './maintenance';

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

function mergeLexicalPlanEntries(
  target: LexicalRemovalPlan['words'],
  source: LexicalRemovalPlan['words'],
): void {
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
      (existing.removedLastSeenMs === null || entry.removedLastSeenMs > existing.removedLastSeenMs)
    ) {
      existing.removedLastSeenMs = entry.removedLastSeenMs;
    }
  }
}

function mergeLexicalPlans(target: LexicalRemovalPlan, source: LexicalRemovalPlan): void {
  mergeLexicalPlanEntries(target.words, source.words);
  mergeLexicalPlanEntries(target.kanji, source.kanji);
}

/**
 * Plan what the delete removes from imm_words/imm_kanji.
 *
 * Deleted videos are planned by video so orphaned subtitle lines (whose session
 * is already gone) still get subtracted; sessions on surviving videos are
 * planned by session. The two scopes are disjoint, so nothing is counted twice.
 */
function planLexicalRemovalsForDelete(
  db: DatabaseSync,
  sessionIdsOnSurvivingVideos: number[],
  videoIds: number[],
): LexicalRemovalPlan {
  const combined: LexicalRemovalPlan = { words: [], kanji: [] };
  forEachIdChunk(sessionIdsOnSurvivingVideos, (chunk) => {
    mergeLexicalPlans(combined, planLexicalRemovalsForSessions(db, chunk));
  });
  forEachIdChunk(videoIds, (chunk) => {
    mergeLexicalPlans(combined, planLexicalRemovalsForVideos(db, chunk));
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
    const sessionIdsOnDeletedVideos = new Set(
      selectIds(
        db,
        (placeholders) => `SELECT session_id FROM imm_sessions WHERE video_id IN (${placeholders})`,
        videoIdList,
        'session_id',
      ),
    );
    for (const sessionId of sessionIdsOnDeletedVideos) sessionIds.add(sessionId);

    const sessionIdList = [...sessionIds];
    const sessionIdsOnSurvivingVideos = sessionIdList.filter(
      (sessionId) => !sessionIdsOnDeletedVideos.has(sessionId),
    );

    // Both plans must be measured before any rows are removed.
    const lexicalRemovals = planLexicalRemovalsForDelete(
      db,
      sessionIdsOnSurvivingVideos,
      videoIdList,
    );
    const lifetimeRemovals = planLifetimeRemovals(db, {
      deletedSessionIds: sessionIdList,
      sessionIdsOnSurvivingVideos,
      deletedVideoIds: videoIdList,
      deletedAnimeIds: animeIdList,
    });
    const affectedRollupGroups: RollupGroup[] = [];
    forEachIdChunk(sessionIdsOnSurvivingVideos, (chunk) => {
      affectedRollupGroups.push(...getRollupGroupsForSessions(db, chunk));
    });
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
    }

    deleteSessionsByIds(db, sessionIdList);
    forEachIdChunk(sessionIdList, (chunk) => {
      const placeholders = makePlaceholders(chunk);
      db.prepare(
        `DELETE FROM imm_lifetime_applied_sessions WHERE session_id IN (${placeholders})`,
      ).run(...chunk);
    });
    forEachIdChunk(videoIdList, (chunk) => {
      const placeholders = makePlaceholders(chunk);
      db.prepare(`DELETE FROM imm_subtitle_lines WHERE video_id IN (${placeholders})`).run(
        ...chunk,
      );
      db.prepare(`DELETE FROM imm_daily_rollups WHERE video_id IN (${placeholders})`).run(...chunk);
      db.prepare(`DELETE FROM imm_monthly_rollups WHERE video_id IN (${placeholders})`).run(
        ...chunk,
      );
      db.prepare(`DELETE FROM imm_media_art WHERE video_id IN (${placeholders})`).run(...chunk);
      db.prepare(`DELETE FROM imm_lifetime_media WHERE video_id IN (${placeholders})`).run(
        ...chunk,
      );
      db.prepare(`DELETE FROM imm_videos WHERE video_id IN (${placeholders})`).run(...chunk);
    });

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
    applyLifetimeRemovals(db, lifetimeRemovals);
    refreshRollupsForGroupsInTransaction(db, affectedRollupGroups);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}
