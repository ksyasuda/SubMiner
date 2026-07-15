import fs from 'node:fs';
import {
  LexiconResolver,
  mergeAnime,
  mergeExcludedWords,
  mergeMediaMetadata,
  mergeVideos,
} from './merge-catalog';
import { mergeSessions } from './merge-sessions';
import { copyRemoteOnlyRollups, refreshRollupsForNewSessions } from './merge-rollups';
import { assertMergeableSchema, createEmptyMergeSummary, type SyncMergeSummary } from './shared';
import { openLibsqlSyncDb, type SyncDb } from './libsql-driver';

export type { SyncMergeSummary } from './shared';

/**
 * Merge a snapshot of another machine's immersion database into the local
 * one. Insert-only union keyed on natural keys (session_uuid, video_key,
 * normalized_title_key, word/kanji identity); lifetime and rollup aggregates
 * are updated incrementally so history older than the session retention
 * window is preserved on both sides. Idempotent: re-merging the same
 * snapshot is a no-op.
 */
export function mergeSnapshotIntoDb(localDbPath: string, snapshotPath: string): SyncMergeSummary {
  if (!fs.existsSync(localDbPath)) {
    throw new Error(`Local stats database not found: ${localDbPath}`);
  }
  if (!fs.existsSync(snapshotPath)) {
    throw new Error(`Snapshot database not found: ${snapshotPath}`);
  }

  const remote = openLibsqlSyncDb(snapshotPath, { readonly: true });
  let local: SyncDb;
  try {
    local = openLibsqlSyncDb(localDbPath, { create: false });
  } catch (error) {
    remote.close();
    throw error;
  }
  try {
    assertMergeableSchema(remote, 'Snapshot');
    assertMergeableSchema(local, 'Local');

    const summary = createEmptyMergeSummary();
    local.exec('PRAGMA foreign_keys = ON');
    local.exec('PRAGMA busy_timeout = 5000');
    local.exec('BEGIN IMMEDIATE');
    try {
      const animeIdMap = mergeAnime(local, remote, summary);
      const { videoIdMap, addedVideoIds } = mergeVideos(local, remote, animeIdMap, summary);
      mergeMediaMetadata(local, remote, videoIdMap, addedVideoIds);
      mergeExcludedWords(local, remote, summary);

      const lexicon = new LexiconResolver(local, remote, summary);
      const { newSessionIds } = mergeSessions(
        local,
        remote,
        videoIdMap,
        animeIdMap,
        lexicon,
        summary,
      );
      lexicon.applyFrequencyDeltas();
      refreshRollupsForNewSessions(local, newSessionIds, summary);
      copyRemoteOnlyRollups(local, remote, videoIdMap, summary);

      local.exec('COMMIT');
      return summary;
    } catch (error) {
      local.exec('ROLLBACK');
      throw error;
    }
  } finally {
    local.close();
    remote.close();
  }
}

export function formatMergeSummary(summary: SyncMergeSummary): string {
  const lines = [
    `Sessions merged: ${summary.sessionsMerged} (${summary.sessionsAlreadyPresent} already present, ${summary.activeSessionsSkipped} unfinished skipped)`,
  ];
  const detail: string[] = [];
  if (summary.animeAdded) detail.push(`${summary.animeAdded} series`);
  if (summary.videosAdded) detail.push(`${summary.videosAdded} videos`);
  if (summary.wordsAdded) detail.push(`${summary.wordsAdded} words`);
  if (summary.kanjiAdded) detail.push(`${summary.kanjiAdded} kanji`);
  if (summary.subtitleLinesAdded) detail.push(`${summary.subtitleLinesAdded} subtitle lines`);
  if (summary.excludedWordsAdded) detail.push(`${summary.excludedWordsAdded} excluded words`);
  if (detail.length > 0) lines.push(`Added: ${detail.join(', ')}`);
  if (summary.dailyRollupsCopied || summary.monthlyRollupsCopied) {
    lines.push(
      `Historical rollups copied: ${summary.dailyRollupsCopied} daily, ${summary.monthlyRollupsCopied} monthly`,
    );
  }
  if (summary.rollupGroupsRecomputed) {
    lines.push(`Rollup groups recomputed: ${summary.rollupGroupsRecomputed}`);
  }
  return lines.join('\n');
}
