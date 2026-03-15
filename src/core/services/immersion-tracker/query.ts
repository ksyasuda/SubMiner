import type { DatabaseSync } from './sqlite';
import type {
  AnimeAnilistEntryRow,
  AnimeDetailRow,
  AnimeEpisodeRow,
  AnimeLibraryRow,
  AnimeWordRow,
  EpisodeCardEventRow,
  EpisodesPerDayRow,
  ImmersionSessionRollupRow,
  KanjiAnimeAppearanceRow,
  KanjiDetailRow,
  KanjiOccurrenceRow,
  KanjiStatsRow,
  KanjiWordRow,
  MediaArtRow,
  MediaDetailRow,
  MediaLibraryRow,
  NewAnimePerDayRow,
  SessionEventRow,
  SessionSummaryQueryRow,
  SessionTimelineRow,
  SimilarWordRow,
  StreakCalendarRow,
  VocabularyCleanupSummary,
  WatchTimePerAnimeRow,
  WordAnimeAppearanceRow,
  WordDetailRow,
  WordOccurrenceRow,
  VocabularyStatsRow,
} from './types';
import { PartOfSpeech, type MergedToken } from '../../../types';
import { shouldExcludeTokenFromVocabularyPersistence } from '../tokenizer/annotation-stage';
import { deriveStoredPartOfSpeech } from '../tokenizer/part-of-speech';

type CleanupVocabularyRow = {
  id: number;
  word: string;
  headword: string;
  reading: string | null;
  part_of_speech: string | null;
  pos1: string | null;
  pos2: string | null;
  pos3: string | null;
  first_seen: number | null;
  last_seen: number | null;
  frequency: number | null;
};

type ResolvedVocabularyPos = {
  headword: string;
  reading: string;
  hasPosMetadata: boolean;
  partOfSpeech: PartOfSpeech;
  pos1: string;
  pos2: string;
  pos3: string;
};

type CleanupVocabularyStatsOptions = {
  resolveLegacyPos?: (row: CleanupVocabularyRow) => Promise<{
    headword: string;
    reading: string;
    partOfSpeech: string;
    pos1: string;
    pos2: string;
    pos3: string;
  } | null>;
};

export function getSessionSummaries(db: DatabaseSync, limit = 50): SessionSummaryQueryRow[] {
  const prepared = db.prepare(`
    SELECT
      s.session_id AS sessionId,
      s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      v.anime_id AS animeId,
      a.canonical_title AS animeTitle,
      s.started_at_ms AS startedAtMs,
      s.ended_at_ms AS endedAtMs,
      COALESCE(MAX(t.total_watched_ms), 0) AS totalWatchedMs,
      COALESCE(MAX(t.active_watched_ms), 0) AS activeWatchedMs,
      COALESCE(MAX(t.lines_seen), 0) AS linesSeen,
      COALESCE(MAX(t.words_seen), 0) AS wordsSeen,
      COALESCE(MAX(t.tokens_seen), 0) AS tokensSeen,
      COALESCE(MAX(t.cards_mined), 0) AS cardsMined,
      COALESCE(MAX(t.lookup_count), 0) AS lookupCount,
      COALESCE(MAX(t.lookup_hits), 0) AS lookupHits
    FROM imm_sessions s
    LEFT JOIN imm_session_telemetry t ON t.session_id = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
    GROUP BY s.session_id
    ORDER BY s.started_at_ms DESC
    LIMIT ?
  `);
  return prepared.all(limit) as unknown as SessionSummaryQueryRow[];
}

export function getSessionTimeline(
  db: DatabaseSync,
  sessionId: number,
  limit = 200,
): SessionTimelineRow[] {
  const prepared = db.prepare(`
    SELECT
      sample_ms AS sampleMs,
      total_watched_ms AS totalWatchedMs,
      active_watched_ms AS activeWatchedMs,
      lines_seen AS linesSeen,
      words_seen AS wordsSeen,
      tokens_seen AS tokensSeen,
      cards_mined AS cardsMined
    FROM imm_session_telemetry
    WHERE session_id = ?
    ORDER BY sample_ms DESC, telemetry_id DESC
    LIMIT ?
  `);
  return prepared.all(sessionId, limit) as unknown as SessionTimelineRow[];
}

export function getQueryHints(db: DatabaseSync): {
  totalSessions: number;
  activeSessions: number;
  episodesToday: number;
  activeAnimeCount: number;
} {
  const sessions = db.prepare('SELECT COUNT(*) AS total FROM imm_sessions');
  const active = db.prepare('SELECT COUNT(*) AS total FROM imm_sessions WHERE ended_at_ms IS NULL');
  const totalSessions = Number((sessions.get() as { total?: number } | null)?.total ?? 0);
  const activeSessions = Number((active.get() as { total?: number } | null)?.total ?? 0);

  const now = new Date();
  const todayLocal = Math.floor(new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime() / 86_400_000);
  const episodesToday = (db.prepare(`
    SELECT COUNT(DISTINCT s.video_id) AS count
    FROM imm_sessions s
    WHERE CAST(julianday(s.started_at_ms / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) = ?
  `).get(todayLocal) as { count: number })?.count ?? 0;

  const thirtyDaysAgoMs = Date.now() - 30 * 86400000;
  const activeAnimeCount = (db.prepare(`
    SELECT COUNT(DISTINCT v.anime_id) AS count
    FROM imm_sessions s
    JOIN imm_videos v ON v.video_id = s.video_id
    WHERE v.anime_id IS NOT NULL
    AND s.started_at_ms >= ?
  `).get(thirtyDaysAgoMs) as { count: number })?.count ?? 0;

  return { totalSessions, activeSessions, episodesToday, activeAnimeCount };
}

export function getDailyRollups(db: DatabaseSync, limit = 60): ImmersionSessionRollupRow[] {
  const prepared = db.prepare(`
    SELECT
      rollup_day AS rollupDayOrMonth,
      video_id AS videoId,
      total_sessions AS totalSessions,
      total_active_min AS totalActiveMin,
      total_lines_seen AS totalLinesSeen,
      total_words_seen AS totalWordsSeen,
      total_tokens_seen AS totalTokensSeen,
      total_cards AS totalCards,
      cards_per_hour AS cardsPerHour,
      words_per_min AS wordsPerMin,
      lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups
    ORDER BY rollup_day DESC, video_id DESC
    LIMIT ?
  `);
  return prepared.all(limit) as unknown as ImmersionSessionRollupRow[];
}

export function getMonthlyRollups(db: DatabaseSync, limit = 24): ImmersionSessionRollupRow[] {
  const prepared = db.prepare(`
    SELECT
      rollup_month AS rollupDayOrMonth,
      video_id AS videoId,
      total_sessions AS totalSessions,
      total_active_min AS totalActiveMin,
      total_lines_seen AS totalLinesSeen,
      total_words_seen AS totalWordsSeen,
      total_tokens_seen AS totalTokensSeen,
      total_cards AS totalCards,
      0 AS cardsPerHour,
      0 AS wordsPerMin,
      0 AS lookupHitRate
    FROM imm_monthly_rollups
    ORDER BY rollup_month DESC, video_id DESC
    LIMIT ?
  `);
  return prepared.all(limit) as unknown as ImmersionSessionRollupRow[];
}

export function getVocabularyStats(
  db: DatabaseSync,
  limit = 100,
  excludePos?: string[],
): VocabularyStatsRow[] {
  const hasExclude = excludePos && excludePos.length > 0;
  const placeholders = hasExclude ? excludePos.map(() => '?').join(', ') : '';
  const whereClause = hasExclude
    ? `WHERE (part_of_speech IS NULL OR part_of_speech NOT IN (${placeholders}))`
    : '';
  const stmt = db.prepare(`
    SELECT id AS wordId, headword, word, reading,
      part_of_speech AS partOfSpeech, pos1, pos2, pos3,
      frequency, first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_words ${whereClause} ORDER BY frequency DESC LIMIT ?
  `);
  const params = hasExclude ? [...excludePos, limit] : [limit];
  return stmt.all(...params) as VocabularyStatsRow[];
}

function toStoredWordToken(row: {
  word: string;
  headword: string;
  part_of_speech: string | null;
  pos1: string | null;
  pos2: string | null;
  pos3: string | null;
}): MergedToken {
  return {
    surface: row.word || row.headword || '',
    reading: '',
    headword: row.headword || row.word || '',
    startPos: 0,
    endPos: 0,
    partOfSpeech: deriveStoredPartOfSpeech({
      partOfSpeech: row.part_of_speech,
      pos1: row.pos1,
    }),
    pos1: row.pos1 ?? '',
    pos2: row.pos2 ?? '',
    pos3: row.pos3 ?? '',
    isMerged: true,
    isKnown: false,
    isNPlusOneTarget: false,
  };
}

function normalizePosField(value: string | null | undefined): string {
  return typeof value === 'string' ? value.trim() : '';
}

function resolveStoredVocabularyPos(row: CleanupVocabularyRow): ResolvedVocabularyPos | null {
  const headword = normalizePosField(row.headword);
  const reading = normalizePosField(row.reading);
  const partOfSpeechRaw = typeof row.part_of_speech === 'string' ? row.part_of_speech.trim() : '';
  const pos1 = normalizePosField(row.pos1);
  const pos2 = normalizePosField(row.pos2);
  const pos3 = normalizePosField(row.pos3);

  if (!headword && !reading && !partOfSpeechRaw && !pos1 && !pos2 && !pos3) {
    return null;
  }

  return {
    headword: headword || normalizePosField(row.word),
    reading,
    hasPosMetadata: Boolean(partOfSpeechRaw || pos1 || pos2 || pos3),
    partOfSpeech: deriveStoredPartOfSpeech({
      partOfSpeech: partOfSpeechRaw,
      pos1,
    }),
    pos1,
    pos2,
    pos3,
  };
}

function hasStructuredPos(pos: ResolvedVocabularyPos | null): boolean {
  return Boolean(pos?.hasPosMetadata && (pos.pos1 || pos.pos2 || pos.pos3 || pos.partOfSpeech));
}

function needsLegacyVocabularyMetadataRepair(
  row: CleanupVocabularyRow,
  stored: ResolvedVocabularyPos | null,
): boolean {
  if (!stored) {
    return true;
  }

  if (!hasStructuredPos(stored)) {
    return true;
  }

  if (!stored.reading) {
    return true;
  }

  if (!stored.headword) {
    return true;
  }

  return stored.headword === normalizePosField(row.word);
}

function shouldUpdateStoredVocabularyPos(
  row: CleanupVocabularyRow,
  next: ResolvedVocabularyPos,
): boolean {
  return (
    normalizePosField(row.headword) !== next.headword ||
    normalizePosField(row.reading) !== next.reading ||
    (next.hasPosMetadata &&
      (normalizePosField(row.part_of_speech) !== next.partOfSpeech ||
        normalizePosField(row.pos1) !== next.pos1 ||
        normalizePosField(row.pos2) !== next.pos2 ||
        normalizePosField(row.pos3) !== next.pos3))
  );
}

function chooseMergedPartOfSpeech(
  current: string | null | undefined,
  incoming: ResolvedVocabularyPos,
): string {
  const normalizedCurrent = normalizePosField(current);
  if (
    normalizedCurrent &&
    normalizedCurrent !== PartOfSpeech.other &&
    incoming.partOfSpeech === PartOfSpeech.other
  ) {
    return normalizedCurrent;
  }
  return incoming.partOfSpeech;
}

async function maybeResolveLegacyVocabularyPos(
  row: CleanupVocabularyRow,
  options: CleanupVocabularyStatsOptions,
): Promise<ResolvedVocabularyPos | null> {
  const stored = resolveStoredVocabularyPos(row);
  if (!needsLegacyVocabularyMetadataRepair(row, stored) || !options.resolveLegacyPos) {
    return stored;
  }

  const resolved = await options.resolveLegacyPos(row);
  if (resolved) {
    return {
      headword: normalizePosField(resolved.headword) || normalizePosField(row.word),
      reading: normalizePosField(resolved.reading),
      hasPosMetadata: true,
      partOfSpeech: deriveStoredPartOfSpeech({
        partOfSpeech: resolved.partOfSpeech,
        pos1: resolved.pos1,
      }),
      pos1: normalizePosField(resolved.pos1),
      pos2: normalizePosField(resolved.pos2),
      pos3: normalizePosField(resolved.pos3),
    };
  }

  return stored;
}

export async function cleanupVocabularyStats(
  db: DatabaseSync,
  options: CleanupVocabularyStatsOptions = {},
): Promise<VocabularyCleanupSummary> {
  const rows = db
    .prepare(
      `SELECT id, word, headword, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
       FROM imm_words`,
    )
    .all() as CleanupVocabularyRow[];
  const findDuplicateStmt = db.prepare(
    `SELECT id, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
     FROM imm_words
     WHERE headword = ? AND word = ? AND reading = ? AND id != ?`,
  );
  const deleteStmt = db.prepare('DELETE FROM imm_words WHERE id = ?');
  const updateStmt = db.prepare(
    `UPDATE imm_words
     SET headword = ?, reading = ?, part_of_speech = ?, pos1 = ?, pos2 = ?, pos3 = ?
     WHERE id = ?`,
  );
  const mergeWordStmt = db.prepare(
    `UPDATE imm_words
     SET
       frequency = COALESCE(frequency, 0) + ?,
       part_of_speech = ?,
       pos1 = ?,
       pos2 = ?,
       pos3 = ?,
       first_seen = MIN(COALESCE(first_seen, ?), ?),
       last_seen = MAX(COALESCE(last_seen, ?), ?)
     WHERE id = ?`,
  );
  const moveOccurrencesStmt = db.prepare(
    `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
     SELECT line_id, ?, occurrence_count
     FROM imm_word_line_occurrences
     WHERE word_id = ?
     ON CONFLICT(line_id, word_id) DO UPDATE SET
       occurrence_count = imm_word_line_occurrences.occurrence_count + excluded.occurrence_count`,
  );
  const deleteOccurrencesStmt = db.prepare('DELETE FROM imm_word_line_occurrences WHERE word_id = ?');
  let kept = 0;
  let deleted = 0;
  let repaired = 0;

  for (const row of rows) {
    const resolvedPos = await maybeResolveLegacyVocabularyPos(row, options);
    const shouldRepair = Boolean(resolvedPos && shouldUpdateStoredVocabularyPos(row, resolvedPos));
    if (resolvedPos && shouldRepair) {
      const duplicate = findDuplicateStmt.get(
        resolvedPos.headword,
        row.word,
        resolvedPos.reading,
        row.id,
      ) as
        | {
            id: number;
            part_of_speech: string | null;
            pos1: string | null;
            pos2: string | null;
            pos3: string | null;
            first_seen: number | null;
            last_seen: number | null;
            frequency: number | null;
          }
        | null;
      if (duplicate) {
        moveOccurrencesStmt.run(duplicate.id, row.id);
        deleteOccurrencesStmt.run(row.id);
        mergeWordStmt.run(
          row.frequency ?? 0,
          chooseMergedPartOfSpeech(duplicate.part_of_speech, resolvedPos),
          normalizePosField(duplicate.pos1) || resolvedPos.pos1,
          normalizePosField(duplicate.pos2) || resolvedPos.pos2,
          normalizePosField(duplicate.pos3) || resolvedPos.pos3,
          row.first_seen ?? duplicate.first_seen ?? 0,
          row.first_seen ?? duplicate.first_seen ?? 0,
          row.last_seen ?? duplicate.last_seen ?? 0,
          row.last_seen ?? duplicate.last_seen ?? 0,
          duplicate.id,
        );
        deleteStmt.run(row.id);
        repaired += 1;
        deleted += 1;
        continue;
      }

      updateStmt.run(
        resolvedPos.headword,
        resolvedPos.reading,
        resolvedPos.partOfSpeech,
        resolvedPos.pos1,
        resolvedPos.pos2,
        resolvedPos.pos3,
        row.id,
      );
      repaired += 1;
    }

    const effectiveRow = {
      ...row,
      headword: resolvedPos?.headword ?? row.headword,
      reading: resolvedPos?.reading ?? row.reading,
      part_of_speech: resolvedPos?.hasPosMetadata ? resolvedPos.partOfSpeech : row.part_of_speech,
      pos1: resolvedPos?.pos1 ?? row.pos1,
      pos2: resolvedPos?.pos2 ?? row.pos2,
      pos3: resolvedPos?.pos3 ?? row.pos3,
    };
    const missingPos =
      !normalizePosField(effectiveRow.part_of_speech) &&
      !normalizePosField(effectiveRow.pos1) &&
      !normalizePosField(effectiveRow.pos2) &&
      !normalizePosField(effectiveRow.pos3);
    if (missingPos || shouldExcludeTokenFromVocabularyPersistence(toStoredWordToken(effectiveRow))) {
      deleteStmt.run(row.id);
      deleted += 1;
      continue;
    }
    kept += 1;
  }

  return {
    scanned: rows.length,
    kept,
    deleted,
    repaired,
  };
}

export function getKanjiStats(db: DatabaseSync, limit = 100): KanjiStatsRow[] {
  const stmt = db.prepare(`
    SELECT id AS kanjiId, kanji, frequency,
      first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_kanji ORDER BY frequency DESC LIMIT ?
  `);
  return stmt.all(limit) as KanjiStatsRow[];
}

export function getWordOccurrences(
  db: DatabaseSync,
  headword: string,
  word: string,
  reading: string,
  limit = 100,
  offset = 0,
): WordOccurrenceRow[] {
  return db
    .prepare(
      `
        SELECT
          l.anime_id AS animeId,
          a.canonical_title AS animeTitle,
          l.video_id AS videoId,
          v.canonical_title AS videoTitle,
          l.session_id AS sessionId,
          l.line_index AS lineIndex,
          l.segment_start_ms AS segmentStartMs,
          l.segment_end_ms AS segmentEndMs,
          l.text AS text,
          o.occurrence_count AS occurrenceCount
        FROM imm_word_line_occurrences o
        JOIN imm_words w ON w.id = o.word_id
        JOIN imm_subtitle_lines l ON l.line_id = o.line_id
        JOIN imm_videos v ON v.video_id = l.video_id
        LEFT JOIN imm_anime a ON a.anime_id = l.anime_id
        WHERE w.headword = ? AND w.word = ? AND w.reading = ?
        ORDER BY l.CREATED_DATE DESC, l.line_id DESC
        LIMIT ?
        OFFSET ?
      `,
    )
    .all(headword, word, reading, limit, offset) as unknown as WordOccurrenceRow[];
}

export function getKanjiOccurrences(
  db: DatabaseSync,
  kanji: string,
  limit = 100,
  offset = 0,
): KanjiOccurrenceRow[] {
  return db
    .prepare(
      `
        SELECT
          l.anime_id AS animeId,
          a.canonical_title AS animeTitle,
          l.video_id AS videoId,
          v.canonical_title AS videoTitle,
          l.session_id AS sessionId,
          l.line_index AS lineIndex,
          l.segment_start_ms AS segmentStartMs,
          l.segment_end_ms AS segmentEndMs,
          l.text AS text,
          o.occurrence_count AS occurrenceCount
        FROM imm_kanji_line_occurrences o
        JOIN imm_kanji k ON k.id = o.kanji_id
        JOIN imm_subtitle_lines l ON l.line_id = o.line_id
        JOIN imm_videos v ON v.video_id = l.video_id
        LEFT JOIN imm_anime a ON a.anime_id = l.anime_id
        WHERE k.kanji = ?
        ORDER BY l.CREATED_DATE DESC, l.line_id DESC
        LIMIT ?
        OFFSET ?
      `,
    )
    .all(kanji, limit, offset) as unknown as KanjiOccurrenceRow[];
}

export function getSessionEvents(
  db: DatabaseSync,
  sessionId: number,
  limit = 500,
): SessionEventRow[] {
  const stmt = db.prepare(`
    SELECT event_type AS eventType, ts_ms AS tsMs, payload_json AS payload
    FROM imm_session_events WHERE session_id = ? ORDER BY ts_ms ASC LIMIT ?
  `);
  return stmt.all(sessionId, limit) as SessionEventRow[];
}

export function getAnimeLibrary(db: DatabaseSync): AnimeLibraryRow[] {
  return db.prepare(`
    SELECT
      a.anime_id AS animeId,
      a.canonical_title AS canonicalTitle,
      a.anilist_id AS anilistId,
      COUNT(DISTINCT s.session_id) AS totalSessions,
      COALESCE(SUM(sm.max_active_ms), 0) AS totalActiveMs,
      COALESCE(SUM(sm.max_cards), 0) AS totalCards,
      COALESCE(SUM(sm.max_words), 0) AS totalWordsSeen,
      COUNT(DISTINCT v.video_id) AS episodeCount,
      a.episodes_total AS episodesTotal,
      MAX(s.started_at_ms) AS lastWatchedMs
    FROM imm_anime a
    JOIN imm_videos v ON v.anime_id = a.anime_id
    JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.cards_mined) AS max_cards,
        MAX(t.words_seen) AS max_words
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm ON sm.session_id = s.session_id
    GROUP BY a.anime_id
    ORDER BY totalActiveMs DESC, lastWatchedMs DESC, canonicalTitle ASC
  `).all() as unknown as AnimeLibraryRow[];
}

export function getAnimeDetail(db: DatabaseSync, animeId: number): AnimeDetailRow | null {
  return db.prepare(`
    SELECT
      a.anime_id AS animeId,
      a.canonical_title AS canonicalTitle,
      a.anilist_id AS anilistId,
      a.title_romaji AS titleRomaji,
      a.title_english AS titleEnglish,
      a.title_native AS titleNative,
      COUNT(DISTINCT s.session_id) AS totalSessions,
      COALESCE(SUM(sm.max_active_ms), 0) AS totalActiveMs,
      COALESCE(SUM(sm.max_cards), 0) AS totalCards,
      COALESCE(SUM(sm.max_words), 0) AS totalWordsSeen,
      COALESCE(SUM(sm.max_lines), 0) AS totalLinesSeen,
      COALESCE(SUM(sm.max_lookups), 0) AS totalLookupCount,
      COALESCE(SUM(sm.max_hits), 0) AS totalLookupHits,
      COUNT(DISTINCT v.video_id) AS episodeCount,
      MAX(s.started_at_ms) AS lastWatchedMs
    FROM imm_anime a
    JOIN imm_videos v ON v.anime_id = a.anime_id
    JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.cards_mined) AS max_cards,
        MAX(t.words_seen) AS max_words,
        MAX(t.lines_seen) AS max_lines,
        MAX(t.lookup_count) AS max_lookups,
        MAX(t.lookup_hits) AS max_hits
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm ON sm.session_id = s.session_id
    WHERE a.anime_id = ?
    GROUP BY a.anime_id
  `).get(animeId) as unknown as AnimeDetailRow | null;
}

export function getAnimeAnilistEntries(db: DatabaseSync, animeId: number): AnimeAnilistEntryRow[] {
  return db.prepare(`
    SELECT DISTINCT
      m.anilist_id AS anilistId,
      m.title_romaji AS titleRomaji,
      m.title_english AS titleEnglish,
      v.parsed_season AS season
    FROM imm_videos v
    JOIN imm_media_art m ON m.video_id = v.video_id
    WHERE v.anime_id = ?
      AND m.anilist_id IS NOT NULL
    ORDER BY v.parsed_season ASC
  `).all(animeId) as unknown as AnimeAnilistEntryRow[];
}

export function getAnimeEpisodes(db: DatabaseSync, animeId: number): AnimeEpisodeRow[] {
  return db.prepare(`
    SELECT
      v.anime_id AS animeId,
      v.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      v.parsed_title AS parsedTitle,
      v.parsed_season AS season,
      v.parsed_episode AS episode,
      v.duration_ms AS durationMs,
      v.watched AS watched,
      COUNT(DISTINCT s.session_id) AS totalSessions,
      COALESCE(SUM(sm.max_active_ms), 0) AS totalActiveMs,
      COALESCE(SUM(sm.max_cards), 0) AS totalCards,
      COALESCE(SUM(sm.max_words), 0) AS totalWordsSeen,
      MAX(s.started_at_ms) AS lastWatchedMs
    FROM imm_videos v
    JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.cards_mined) AS max_cards,
        MAX(t.words_seen) AS max_words
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm ON sm.session_id = s.session_id
    WHERE v.anime_id = ?
    GROUP BY v.video_id
    ORDER BY
      CASE WHEN v.parsed_season IS NULL THEN 1 ELSE 0 END,
      v.parsed_season ASC,
      CASE WHEN v.parsed_episode IS NULL THEN 1 ELSE 0 END,
      v.parsed_episode ASC,
      v.video_id ASC
  `).all(animeId) as unknown as AnimeEpisodeRow[];
}

export function getMediaLibrary(db: DatabaseSync): MediaLibraryRow[] {
  return db.prepare(`
    SELECT
      v.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      COUNT(DISTINCT s.session_id) AS totalSessions,
      COALESCE(SUM(sm.max_active_ms), 0) AS totalActiveMs,
      COALESCE(SUM(sm.max_cards), 0) AS totalCards,
      COALESCE(SUM(sm.max_words), 0) AS totalWordsSeen,
      MAX(s.started_at_ms) AS lastWatchedMs,
      CASE WHEN ma.cover_blob IS NOT NULL THEN 1 ELSE 0 END AS hasCoverArt
    FROM imm_videos v
    JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.cards_mined) AS max_cards,
        MAX(t.words_seen) AS max_words
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm ON sm.session_id = s.session_id
    LEFT JOIN imm_media_art ma ON ma.video_id = v.video_id
    GROUP BY v.video_id
    ORDER BY lastWatchedMs DESC
  `).all() as unknown as MediaLibraryRow[];
}

export function getMediaDetail(db: DatabaseSync, videoId: number): MediaDetailRow | null {
  return db.prepare(`
    SELECT
      v.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      COUNT(DISTINCT s.session_id) AS totalSessions,
      COALESCE(SUM(sm.max_active_ms), 0) AS totalActiveMs,
      COALESCE(SUM(sm.max_cards), 0) AS totalCards,
      COALESCE(SUM(sm.max_words), 0) AS totalWordsSeen,
      COALESCE(SUM(sm.max_lines), 0) AS totalLinesSeen,
      COALESCE(SUM(sm.max_lookups), 0) AS totalLookupCount,
      COALESCE(SUM(sm.max_hits), 0) AS totalLookupHits
    FROM imm_videos v
    JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.cards_mined) AS max_cards,
        MAX(t.words_seen) AS max_words,
        MAX(t.lines_seen) AS max_lines,
        MAX(t.lookup_count) AS max_lookups,
        MAX(t.lookup_hits) AS max_hits
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm ON sm.session_id = s.session_id
    WHERE v.video_id = ?
    GROUP BY v.video_id
  `).get(videoId) as unknown as MediaDetailRow | null;
}

export function getMediaSessions(db: DatabaseSync, videoId: number, limit = 100): SessionSummaryQueryRow[] {
  return db.prepare(`
    SELECT
      s.session_id AS sessionId,
      s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      s.started_at_ms AS startedAtMs,
      s.ended_at_ms AS endedAtMs,
      COALESCE(MAX(t.total_watched_ms), 0) AS totalWatchedMs,
      COALESCE(MAX(t.active_watched_ms), 0) AS activeWatchedMs,
      COALESCE(MAX(t.lines_seen), 0) AS linesSeen,
      COALESCE(MAX(t.words_seen), 0) AS wordsSeen,
      COALESCE(MAX(t.tokens_seen), 0) AS tokensSeen,
      COALESCE(MAX(t.cards_mined), 0) AS cardsMined,
      COALESCE(MAX(t.lookup_count), 0) AS lookupCount,
      COALESCE(MAX(t.lookup_hits), 0) AS lookupHits
    FROM imm_sessions s
    LEFT JOIN imm_session_telemetry t ON t.session_id = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    WHERE s.video_id = ?
    GROUP BY s.session_id
    ORDER BY s.started_at_ms DESC
    LIMIT ?
  `).all(videoId, limit) as unknown as SessionSummaryQueryRow[];
}

export function getMediaDailyRollups(db: DatabaseSync, videoId: number, limit = 90): ImmersionSessionRollupRow[] {
  return db.prepare(`
    SELECT
      rollup_day AS rollupDayOrMonth,
      video_id AS videoId,
      total_sessions AS totalSessions,
      total_active_min AS totalActiveMin,
      total_lines_seen AS totalLinesSeen,
      total_words_seen AS totalWordsSeen,
      total_tokens_seen AS totalTokensSeen,
      total_cards AS totalCards,
      cards_per_hour AS cardsPerHour,
      words_per_min AS wordsPerMin,
      lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups
    WHERE video_id = ?
    ORDER BY rollup_day DESC
    LIMIT ?
  `).all(videoId, limit) as unknown as ImmersionSessionRollupRow[];
}

export function getAnimeCoverArt(db: DatabaseSync, animeId: number): MediaArtRow | null {
  return db.prepare(`
    SELECT
      a.video_id AS videoId,
      a.anilist_id AS anilistId,
      a.cover_url AS coverUrl,
      a.cover_blob AS coverBlob,
      a.title_romaji AS titleRomaji,
      a.title_english AS titleEnglish,
      a.episodes_total AS episodesTotal,
      a.fetched_at_ms AS fetchedAtMs
    FROM imm_media_art a
    JOIN imm_videos v ON v.video_id = a.video_id
    WHERE v.anime_id = ?
    AND a.cover_blob IS NOT NULL
    LIMIT 1
  `).get(animeId) as unknown as MediaArtRow | null;
}

export function getCoverArt(db: DatabaseSync, videoId: number): MediaArtRow | null {
  return db.prepare(`
    SELECT
      video_id AS videoId,
      anilist_id AS anilistId,
      cover_url AS coverUrl,
      cover_blob AS coverBlob,
      title_romaji AS titleRomaji,
      title_english AS titleEnglish,
      episodes_total AS episodesTotal,
      fetched_at_ms AS fetchedAtMs
    FROM imm_media_art
    WHERE video_id = ?
  `).get(videoId) as unknown as MediaArtRow | null;
}

export function getStreakCalendar(db: DatabaseSync, days = 90): StreakCalendarRow[] {
  const now = new Date();
  const localMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const todayLocalDay = Math.floor(localMidnight / 86_400_000);
  const cutoffDay = todayLocalDay - days;
  return db.prepare(`
    SELECT rollup_day AS epochDay, SUM(total_active_min) AS totalActiveMin
    FROM imm_daily_rollups
    WHERE rollup_day >= ?
    GROUP BY rollup_day
    ORDER BY rollup_day ASC
  `).all(cutoffDay) as StreakCalendarRow[];
}

export function getAnimeWords(db: DatabaseSync, animeId: number, limit = 50): AnimeWordRow[] {
  return db.prepare(`
    SELECT w.id AS wordId, w.headword, w.word, w.reading, w.part_of_speech AS partOfSpeech,
           SUM(o.occurrence_count) AS frequency
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_words w ON w.id = o.word_id
    WHERE sl.anime_id = ?
    GROUP BY w.id
    ORDER BY frequency DESC
    LIMIT ?
  `).all(animeId, limit) as unknown as AnimeWordRow[];
}

export function getAnimeDailyRollups(db: DatabaseSync, animeId: number, limit = 90): ImmersionSessionRollupRow[] {
  return db.prepare(`
    SELECT r.rollup_day AS rollupDayOrMonth, r.video_id AS videoId,
           r.total_sessions AS totalSessions, r.total_active_min AS totalActiveMin,
           r.total_lines_seen AS totalLinesSeen, r.total_words_seen AS totalWordsSeen,
           r.total_tokens_seen AS totalTokensSeen, r.total_cards AS totalCards,
           r.cards_per_hour AS cardsPerHour, r.words_per_min AS wordsPerMin,
           r.lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups r
    JOIN imm_videos v ON v.video_id = r.video_id
    WHERE v.anime_id = ?
    ORDER BY r.rollup_day DESC
    LIMIT ?
  `).all(animeId, limit) as unknown as ImmersionSessionRollupRow[];
}

export function getEpisodesPerDay(db: DatabaseSync, limit = 90): EpisodesPerDayRow[] {
  return db.prepare(`
    SELECT CAST(julianday(s.started_at_ms / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) AS epochDay,
           COUNT(DISTINCT s.video_id) AS episodeCount
    FROM imm_sessions s
    GROUP BY epochDay
    ORDER BY epochDay DESC
    LIMIT ?
  `).all(limit) as EpisodesPerDayRow[];
}

export function getNewAnimePerDay(db: DatabaseSync, limit = 90): NewAnimePerDayRow[] {
  return db.prepare(`
    SELECT first_day AS epochDay, COUNT(*) AS newAnimeCount
    FROM (
      SELECT CAST(julianday(MIN(s.started_at_ms) / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) AS first_day
      FROM imm_sessions s
      JOIN imm_videos v ON v.video_id = s.video_id
      WHERE v.anime_id IS NOT NULL
      GROUP BY v.anime_id
    )
    GROUP BY first_day
    ORDER BY first_day DESC
    LIMIT ?
  `).all(limit) as NewAnimePerDayRow[];
}

export function getWatchTimePerAnime(db: DatabaseSync, limit = 90): WatchTimePerAnimeRow[] {
  const nowD = new Date();
  const cutoffDay = Math.floor(new Date(nowD.getFullYear(), nowD.getMonth(), nowD.getDate()).getTime() / 86_400_000) - limit;
  return db.prepare(`
    SELECT r.rollup_day AS epochDay, a.anime_id AS animeId,
           a.canonical_title AS animeTitle,
           SUM(r.total_active_min) AS totalActiveMin
    FROM imm_daily_rollups r
    JOIN imm_videos v ON v.video_id = r.video_id
    JOIN imm_anime a ON a.anime_id = v.anime_id
    WHERE r.rollup_day >= ?
    GROUP BY r.rollup_day, a.anime_id
    ORDER BY r.rollup_day ASC
  `).all(cutoffDay) as WatchTimePerAnimeRow[];
}

export function getWordDetail(db: DatabaseSync, wordId: number): WordDetailRow | null {
  return db.prepare(`
    SELECT id AS wordId, headword, word, reading,
           part_of_speech AS partOfSpeech, pos1, pos2, pos3,
           frequency, first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_words WHERE id = ?
  `).get(wordId) as WordDetailRow | null;
}

export function getWordAnimeAppearances(db: DatabaseSync, wordId: number): WordAnimeAppearanceRow[] {
  return db.prepare(`
    SELECT a.anime_id AS animeId, a.canonical_title AS animeTitle,
           SUM(o.occurrence_count) AS occurrenceCount
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_anime a ON a.anime_id = sl.anime_id
    WHERE o.word_id = ? AND sl.anime_id IS NOT NULL
    GROUP BY a.anime_id
    ORDER BY occurrenceCount DESC
  `).all(wordId) as WordAnimeAppearanceRow[];
}

export function getSimilarWords(db: DatabaseSync, wordId: number, limit = 10): SimilarWordRow[] {
  const word = db.prepare('SELECT headword, reading FROM imm_words WHERE id = ?').get(wordId) as { headword: string; reading: string } | null;
  if (!word) return [];
  return db.prepare(`
    SELECT id AS wordId, headword, word, reading, frequency
    FROM imm_words
    WHERE id != ?
    AND (reading = ? OR headword LIKE ? OR headword LIKE ?)
    ORDER BY frequency DESC
    LIMIT ?
  `).all(
    wordId,
    word.reading,
    `%${word.headword.charAt(0)}%`,
    `%${word.headword.charAt(word.headword.length - 1)}%`,
    limit,
  ) as SimilarWordRow[];
}

export function getKanjiDetail(db: DatabaseSync, kanjiId: number): KanjiDetailRow | null {
  return db.prepare(`
    SELECT id AS kanjiId, kanji, frequency, first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_kanji WHERE id = ?
  `).get(kanjiId) as KanjiDetailRow | null;
}

export function getKanjiAnimeAppearances(db: DatabaseSync, kanjiId: number): KanjiAnimeAppearanceRow[] {
  return db.prepare(`
    SELECT a.anime_id AS animeId, a.canonical_title AS animeTitle,
           SUM(o.occurrence_count) AS occurrenceCount
    FROM imm_kanji_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_anime a ON a.anime_id = sl.anime_id
    WHERE o.kanji_id = ? AND sl.anime_id IS NOT NULL
    GROUP BY a.anime_id
    ORDER BY occurrenceCount DESC
  `).all(kanjiId) as KanjiAnimeAppearanceRow[];
}

export function getKanjiWords(db: DatabaseSync, kanjiId: number, limit = 20): KanjiWordRow[] {
  const kanjiRow = db.prepare('SELECT kanji FROM imm_kanji WHERE id = ?').get(kanjiId) as { kanji: string } | null;
  if (!kanjiRow) return [];
  return db.prepare(`
    SELECT id AS wordId, headword, word, reading, frequency
    FROM imm_words
    WHERE headword LIKE ?
    ORDER BY frequency DESC
    LIMIT ?
  `).all(`%${kanjiRow.kanji}%`, limit) as KanjiWordRow[];
}

export function getEpisodeWords(db: DatabaseSync, videoId: number, limit = 50): AnimeWordRow[] {
  return db.prepare(`
    SELECT w.id AS wordId, w.headword, w.word, w.reading, w.part_of_speech AS partOfSpeech,
           SUM(o.occurrence_count) AS frequency
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_words w ON w.id = o.word_id
    WHERE sl.video_id = ?
    GROUP BY w.id
    ORDER BY frequency DESC
    LIMIT ?
  `).all(videoId, limit) as unknown as AnimeWordRow[];
}

export function getEpisodeSessions(db: DatabaseSync, videoId: number): SessionSummaryQueryRow[] {
  return db.prepare(`
    SELECT
      s.session_id AS sessionId, s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      s.started_at_ms AS startedAtMs, s.ended_at_ms AS endedAtMs,
      COALESCE(MAX(t.total_watched_ms), 0) AS totalWatchedMs,
      COALESCE(MAX(t.active_watched_ms), 0) AS activeWatchedMs,
      COALESCE(MAX(t.lines_seen), 0) AS linesSeen,
      COALESCE(MAX(t.words_seen), 0) AS wordsSeen,
      COALESCE(MAX(t.tokens_seen), 0) AS tokensSeen,
      COALESCE(MAX(t.cards_mined), 0) AS cardsMined,
      COALESCE(MAX(t.lookup_count), 0) AS lookupCount,
      COALESCE(MAX(t.lookup_hits), 0) AS lookupHits
    FROM imm_sessions s
    JOIN imm_videos v ON v.video_id = s.video_id
    LEFT JOIN imm_session_telemetry t ON t.session_id = s.session_id
    WHERE s.video_id = ?
    GROUP BY s.session_id
    ORDER BY s.started_at_ms DESC
  `).all(videoId) as SessionSummaryQueryRow[];
}

export function getEpisodeCardEvents(db: DatabaseSync, videoId: number): EpisodeCardEventRow[] {
  const rows = db.prepare(`
    SELECT e.event_id AS eventId, e.session_id AS sessionId,
           e.ts_ms AS tsMs, e.cards_delta AS cardsDelta,
           e.payload_json AS payloadJson
    FROM imm_session_events e
    JOIN imm_sessions s ON s.session_id = e.session_id
    WHERE s.video_id = ? AND e.event_type = 4
    ORDER BY e.ts_ms DESC
  `).all(videoId) as Array<{ eventId: number; sessionId: number; tsMs: number; cardsDelta: number; payloadJson: string | null }>;

  return rows.map(row => {
    let noteIds: number[] = [];
    if (row.payloadJson) {
      try {
        const parsed = JSON.parse(row.payloadJson);
        if (Array.isArray(parsed.noteIds)) noteIds = parsed.noteIds;
      } catch {}
    }
    return { eventId: row.eventId, sessionId: row.sessionId, tsMs: row.tsMs, cardsDelta: row.cardsDelta, noteIds };
  });
}

export function upsertCoverArt(
  db: DatabaseSync,
  videoId: number,
  art: {
    anilistId: number | null;
    coverUrl: string | null;
    coverBlob: Buffer | null;
    titleRomaji: string | null;
    titleEnglish: string | null;
    episodesTotal: number | null;
  },
): void {
  const nowMs = Date.now();
  db.prepare(`
    INSERT INTO imm_media_art (
      video_id, anilist_id, cover_url, cover_blob,
      title_romaji, title_english, episodes_total,
      fetched_at_ms, CREATED_DATE, LAST_UPDATE_DATE
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(video_id) DO UPDATE SET
      anilist_id = excluded.anilist_id,
      cover_url = excluded.cover_url,
      cover_blob = excluded.cover_blob,
      title_romaji = excluded.title_romaji,
      title_english = excluded.title_english,
      episodes_total = excluded.episodes_total,
      fetched_at_ms = excluded.fetched_at_ms,
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `).run(
    videoId, art.anilistId, art.coverUrl, art.coverBlob,
    art.titleRomaji, art.titleEnglish, art.episodesTotal,
    nowMs, nowMs, nowMs,
  );
}

export function updateAnimeAnilistInfo(
  db: DatabaseSync,
  videoId: number,
  info: {
    anilistId: number;
    titleRomaji: string | null;
    titleEnglish: string | null;
    titleNative: string | null;
    episodesTotal: number | null;
  },
): void {
  const row = db.prepare('SELECT anime_id FROM imm_videos WHERE video_id = ?').get(videoId) as {
    anime_id: number | null;
  } | null;
  if (!row?.anime_id) return;

  db.prepare(`
    UPDATE imm_anime
    SET
      anilist_id = COALESCE(?, anilist_id),
      title_romaji = COALESCE(?, title_romaji),
      title_english = COALESCE(?, title_english),
      title_native = COALESCE(?, title_native),
      episodes_total = COALESCE(?, episodes_total),
      LAST_UPDATE_DATE = ?
    WHERE anime_id = ?
  `).run(
    info.anilistId,
    info.titleRomaji,
    info.titleEnglish,
    info.titleNative,
    info.episodesTotal,
    Date.now(),
    row.anime_id,
  );
}

export function markVideoWatched(db: DatabaseSync, videoId: number, watched: boolean): void {
  db.prepare('UPDATE imm_videos SET watched = ?, LAST_UPDATE_DATE = ? WHERE video_id = ?')
    .run(watched ? 1 : 0, Date.now(), videoId);
}

export function getVideoDurationMs(db: DatabaseSync, videoId: number): number {
  const row = db.prepare('SELECT duration_ms FROM imm_videos WHERE video_id = ?').get(videoId) as {
    duration_ms: number;
  } | null;
  return row?.duration_ms ?? 0;
}

export function isVideoWatched(db: DatabaseSync, videoId: number): boolean {
  const row = db.prepare('SELECT watched FROM imm_videos WHERE video_id = ?').get(videoId) as {
    watched: number;
  } | null;
  return row?.watched === 1;
}
