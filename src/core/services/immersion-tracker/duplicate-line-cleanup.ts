/*
 * Retroactive removal of animation-burst subtitle lines from the stats database.
 *
 * Before the live ingest gate existed, a karaoke OP recorded one line -- and one count
 * for every word in it -- per animation frame, which is enough to put an OP lyric at the
 * top of "Top Repeated Words" for good. This module finds those runs in what is already
 * stored and takes them back down to one line.
 *
 * Only timing is available here: the stored text has been stripped of ASS markup, so the
 * authoring evidence the file-level parser uses (`\t`, `\move`, karaoke timing, a
 * changing override signature) is long gone. What is left is a run of identical,
 * contiguous, short-lived lines inside a single session.
 *
 * The run has to be as long as the timing-only rule in `subtitle-cue-dedup` demands, but
 * its short frames may be as long as the animation-frame bound rather than the much
 * tighter timing-only one. A qualifying run may end with one longer hold, which is a
 * common karaoke shape. Five or more repeats of the same text, each ending where the next
 * begins, is already conclusive on its own -- no dialogue does that -- and the tighter
 * bound would walk straight past the heavier typesetting that motivated this, where
 * frames sit nearer a quarter of a second. Both bounds are options, so a cautious run can
 * ask for more, and a dry run always reports before anything is removed.
 *
 * Scope: subtitle lines, their word/kanji occurrences, and the `imm_words`/`imm_kanji`
 * aggregates those occurrences feed. Session telemetry (`lines_seen`, `tokens_seen`) and
 * the rollups derived from it are left alone; they are cumulative samples taken at record
 * time, and for sessions whose raw rows have since been pruned they cannot be recomputed.
 */

import type { DatabaseSync } from './sqlite';
import {
  ANIMATION_FRAME_MAX_SECONDS,
  DUPLICATE_CUE_GAP_TOLERANCE_SECONDS,
  MIN_STREAM_RESIDUE_FRAMES,
  MIN_TIMING_ONLY_FRAMES,
  TIMING_ONLY_FRAME_MAX_SECONDS,
} from '../subtitle-burst-constants';
import {
  applyLexicalRemovals,
  makePlaceholders,
  planLexicalRemovalsForLines,
  toDbTimestamp,
} from './query-shared';
import { nowMs } from './time';

const MS_PER_DAY = 86_400_000;
/** SQLite caps bound parameters per statement; stay well under it. */
const ID_BATCH_SIZE = 400;
const DEFAULT_SAMPLE_LIMIT = 20;

export interface DuplicateSubtitleLineCleanupOptions {
  /** Only consider lines recorded within this many days. Null or omitted = all history. */
  lookbackDays?: number | null;
  /** Measure without writing. */
  dryRun?: boolean;
  /** Identical contiguous lines needed before a run counts as an animation. */
  minRunLength?: number;
  /** Longest a single event may last and still look like an animation frame. */
  maxFrameSeconds?: number;
  /** How many of the largest runs to describe in the summary. */
  sampleLimit?: number;
}

export interface DuplicateSubtitleLineBurst {
  sessionId: number;
  videoId: number;
  text: string;
  /** Kept line, extended to cover the whole run. */
  keptLineId: number;
  removedLineIds: number[];
  startMs: number;
  endMs: number;
}

export interface DuplicateSubtitleLineSample {
  videoId: number;
  videoTitle: string | null;
  text: string;
  frames: number;
  removedLines: number;
  startMs: number;
  endMs: number;
}

export interface DuplicateSubtitleLineCleanupSummary {
  dryRun: boolean;
  lookbackDays: number | null;
  scannedLines: number;
  burstGroups: number;
  removedLines: number;
  removedWordOccurrences: number;
  removedKanjiOccurrences: number;
  samples: DuplicateSubtitleLineSample[];
}

export interface StoredSubtitleLineRow {
  lineId: number;
  sessionId: number;
  videoId: number;
  text: string;
  startMs: number;
  endMs: number;
}

interface ResolvedBounds {
  lookbackDays: number | null;
  minRunLength: number;
  maxFrameMs: number;
  /** Shorter runs qualify only when every event sits under this much stricter bound. */
  residueMinRunLength: number;
  strictFrameMs: number;
  gapToleranceMs: number;
  sampleLimit: number;
}

function resolveBounds(options: DuplicateSubtitleLineCleanupOptions): ResolvedBounds {
  const lookbackDays =
    typeof options.lookbackDays === 'number' && Number.isFinite(options.lookbackDays)
      ? Math.max(1, Math.floor(options.lookbackDays))
      : null;
  const minRunLength =
    typeof options.minRunLength === 'number' && Number.isFinite(options.minRunLength)
      ? Math.max(2, Math.floor(options.minRunLength))
      : MIN_TIMING_ONLY_FRAMES;
  const maxFrameSeconds =
    typeof options.maxFrameSeconds === 'number' &&
    Number.isFinite(options.maxFrameSeconds) &&
    options.maxFrameSeconds > 0
      ? options.maxFrameSeconds
      : ANIMATION_FRAME_MAX_SECONDS;
  const sampleLimit =
    typeof options.sampleLimit === 'number' && options.sampleLimit >= 0
      ? Math.floor(options.sampleLimit)
      : DEFAULT_SAMPLE_LIMIT;
  return {
    lookbackDays,
    minRunLength,
    maxFrameMs: Math.round(maxFrameSeconds * 1000),
    residueMinRunLength: Math.max(MIN_STREAM_RESIDUE_FRAMES, minRunLength - 1),
    strictFrameMs: Math.round(TIMING_ONLY_FRAME_MAX_SECONDS * 1000),
    gapToleranceMs: Math.round(DUPLICATE_CUE_GAP_TOLERANCE_SECONDS * 1000),
    sampleLimit,
  };
}

/**
 * `CREATED_DATE` holds epoch milliseconds on rows this app wrote, but older and synced
 * rows can carry seconds, so normalize before comparing against the cutoff.
 */
const CREATED_MS_SQL = `
  CASE
    WHEN sl.CREATED_DATE < 10000000000 THEN sl.CREATED_DATE * 1000
    ELSE sl.CREATED_DATE
  END`;

function readCandidateLines(db: DatabaseSync, bounds: ResolvedBounds): StoredSubtitleLineRow[] {
  const scope =
    bounds.lookbackDays === null
      ? ''
      : `AND sl.CREATED_DATE IS NOT NULL AND ${CREATED_MS_SQL} >= ?`;
  const params = bounds.lookbackDays === null ? [] : [nowMs() - bounds.lookbackDays * MS_PER_DAY];
  return db
    .prepare(
      `SELECT
         sl.line_id AS lineId,
         sl.session_id AS sessionId,
         sl.video_id AS videoId,
         sl.text AS text,
         sl.segment_start_ms AS startMs,
         sl.segment_end_ms AS endMs
       FROM imm_subtitle_lines sl
       WHERE sl.segment_start_ms IS NOT NULL
         AND sl.segment_end_ms IS NOT NULL
         ${scope}
       ORDER BY sl.session_id, sl.video_id, sl.segment_start_ms, sl.line_id`,
    )
    .all(...params) as StoredSubtitleLineRow[];
}

function isBurst(run: StoredSubtitleLineRow[], bounds: ResolvedBounds): boolean {
  const isShortFrame = (row: StoredSubtitleLineRow): boolean =>
    row.endMs - row.startMs <= bounds.maxFrameMs;
  // The residue the live gate leaves behind: it records the first frames of a burst
  // before the run is long enough to recognise, so one frame fewer than the timing-only
  // minimum, every one under the strict timing-only bound. No dialogue holds identical
  // sub-tenth-second lines back to back that many times.
  if (
    run.length >= bounds.residueMinRunLength &&
    run.every((row) => row.endMs - row.startMs <= bounds.strictFrameMs)
  ) {
    return true;
  }
  if (run.length < bounds.minRunLength) {
    return false;
  }
  if (run.every(isShortFrame)) {
    return true;
  }
  // Karaoke commonly finishes its short animation frames with one long hold. Only the
  // final event may exceed the frame bound, and the short frames before it must already
  // meet the minimum run length on their own.
  return (
    run.length - 1 >= bounds.minRunLength &&
    run.slice(0, -1).every(isShortFrame) &&
    !isShortFrame(run[run.length - 1]!)
  );
}

function toBurst(run: StoredSubtitleLineRow[]): DuplicateSubtitleLineBurst {
  const [first] = run;
  return {
    sessionId: first!.sessionId,
    videoId: first!.videoId,
    text: first!.text,
    keptLineId: first!.lineId,
    removedLineIds: run.slice(1).map((row) => row.lineId),
    startMs: first!.startMs,
    endMs: run.reduce((latest, row) => Math.max(latest, row.endMs), first!.endMs),
  };
}

/**
 * Group stored lines into animation runs.
 *
 * Rows are bucketed per (session, video, text) before chaining, the way the file-level
 * dedup buckets cues: dual-line karaoke interleaves two texts frame by frame, and
 * chaining across the interleave would break every run at length one.
 *
 * Runs never cross a session, which is what keeps a rewatch intact: the same episode
 * watched twice stores the same line twice, and those two belong to different sessions.
 */
export function findDuplicateSubtitleLineBursts(
  rows: readonly StoredSubtitleLineRow[],
  options: DuplicateSubtitleLineCleanupOptions = {},
): DuplicateSubtitleLineBurst[] {
  const bounds = resolveBounds(options);

  // Insertion order preserves the query's startMs ordering within each bucket.
  const rowsByKey = new Map<string, StoredSubtitleLineRow[]>();
  for (const row of rows) {
    const key = `${row.sessionId}|${row.videoId}|${row.text}`;
    const bucket = rowsByKey.get(key);
    if (bucket) {
      bucket.push(row);
    } else {
      rowsByKey.set(key, [row]);
    }
  }

  const bursts: DuplicateSubtitleLineBurst[] = [];
  for (const bucket of rowsByKey.values()) {
    if (bucket.length < 2) {
      continue;
    }

    let run: StoredSubtitleLineRow[] = [];
    let chainEndMs = 0;

    const closeRun = (): void => {
      if (run.length > 1 && isBurst(run, bounds)) {
        bursts.push(toBurst(run));
      }
      run = [];
    };

    for (const row of bucket) {
      if (run.length > 0 && row.startMs <= chainEndMs + bounds.gapToleranceMs) {
        run.push(row);
        chainEndMs = Math.max(chainEndMs, row.endMs);
        continue;
      }
      closeRun();
      run = [row];
      chainEndMs = row.endMs;
    }
    closeRun();
  }

  return bursts;
}

function chunk<T>(values: T[], size: number): T[][] {
  const chunks: T[][] = [];
  for (let i = 0; i < values.length; i += size) {
    chunks.push(values.slice(i, i + size));
  }
  return chunks;
}

function buildSamples(
  db: DatabaseSync,
  bursts: DuplicateSubtitleLineBurst[],
  sampleLimit: number,
): DuplicateSubtitleLineSample[] {
  if (sampleLimit === 0 || bursts.length === 0) {
    return [];
  }
  const largest = [...bursts]
    .sort((a, b) => b.removedLineIds.length - a.removedLineIds.length)
    .slice(0, sampleLimit);
  const videoIds = [...new Set(largest.map((burst) => burst.videoId))];
  const titles = new Map<number, string>();
  for (const batch of chunk(videoIds, ID_BATCH_SIZE)) {
    const rows = db
      .prepare(
        `SELECT video_id AS videoId, canonical_title AS title
         FROM imm_videos
         WHERE video_id IN (${makePlaceholders(batch)})`,
      )
      .all(...batch) as Array<{ videoId: number; title: string | null }>;
    for (const row of rows) {
      if (row.title) titles.set(row.videoId, row.title);
    }
  }

  return largest.map((burst) => ({
    videoId: burst.videoId,
    videoTitle: titles.get(burst.videoId) ?? null,
    text: burst.text,
    frames: burst.removedLineIds.length + 1,
    removedLines: burst.removedLineIds.length,
    startMs: burst.startMs,
    endMs: burst.endMs,
  }));
}

function sumRemovedOccurrences(
  db: DatabaseSync,
  table: 'imm_word_line_occurrences' | 'imm_kanji_line_occurrences',
  lineIds: number[],
): number {
  let total = 0;
  for (const batch of chunk(lineIds, ID_BATCH_SIZE)) {
    const row = db
      .prepare(
        `SELECT COALESCE(SUM(occurrence_count), 0) AS total
         FROM ${table}
         WHERE line_id IN (${makePlaceholders(batch)})`,
      )
      .get(...batch) as { total: number } | null;
    total += row?.total ?? 0;
  }
  return total;
}

function applyBursts(db: DatabaseSync, bursts: DuplicateSubtitleLineBurst[]): void {
  const removedLineIds = bursts.flatMap((burst) => burst.removedLineIds);
  const currentMs = toDbTimestamp(nowMs());

  db.exec('BEGIN IMMEDIATE');
  try {
    for (const batch of chunk(removedLineIds, ID_BATCH_SIZE)) {
      const placeholders = makePlaceholders(batch);
      // Measured before the delete, applied after it: `applyLexicalRemovals` checks the
      // surviving occurrences to decide whether a zeroed count really means the word is
      // gone, so the rows it inspects have to be the post-delete ones.
      const plan = planLexicalRemovalsForLines(db, batch);
      db.prepare(`DELETE FROM imm_word_line_occurrences WHERE line_id IN (${placeholders})`).run(
        ...batch,
      );
      db.prepare(`DELETE FROM imm_kanji_line_occurrences WHERE line_id IN (${placeholders})`).run(
        ...batch,
      );
      db.prepare(`DELETE FROM imm_subtitle_lines WHERE line_id IN (${placeholders})`).run(...batch);
      applyLexicalRemovals(db, plan);
    }

    const extendStmt = db.prepare(
      `UPDATE imm_subtitle_lines
       SET segment_end_ms = ?, LAST_UPDATE_DATE = ?
       WHERE line_id = ? AND (segment_end_ms IS NULL OR segment_end_ms < ?)`,
    );
    for (const burst of bursts) {
      extendStmt.run(burst.endMs, currentMs, burst.keptLineId, burst.endMs);
    }
    db.exec('COMMIT');
  } catch (error) {
    try {
      db.exec('ROLLBACK');
    } catch {
      // Surface the transaction failure, not the rollback's.
    }
    throw error;
  }
}

/**
 * Collapse stored animation bursts down to one line each.
 *
 * A dry run measures exactly what an apply would remove, using the same scan, so the
 * numbers shown in a confirmation prompt are the numbers that will happen.
 */
export function cleanupDuplicateSubtitleLines(
  db: DatabaseSync,
  options: DuplicateSubtitleLineCleanupOptions = {},
): DuplicateSubtitleLineCleanupSummary {
  const bounds = resolveBounds(options);
  const dryRun = options.dryRun === true;
  const rows = readCandidateLines(db, bounds);
  const bursts = findDuplicateSubtitleLineBursts(rows, options);
  const removedLineIds = bursts.flatMap((burst) => burst.removedLineIds);

  const summary: DuplicateSubtitleLineCleanupSummary = {
    dryRun,
    lookbackDays: bounds.lookbackDays,
    scannedLines: rows.length,
    burstGroups: bursts.length,
    removedLines: removedLineIds.length,
    removedWordOccurrences: sumRemovedOccurrences(db, 'imm_word_line_occurrences', removedLineIds),
    removedKanjiOccurrences: sumRemovedOccurrences(
      db,
      'imm_kanji_line_occurrences',
      removedLineIds,
    ),
    samples: buildSamples(db, bursts, bounds.sampleLimit),
  };

  if (dryRun || removedLineIds.length === 0) {
    return summary;
  }

  applyBursts(db, bursts);
  return summary;
}
