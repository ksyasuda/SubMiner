import type { DatabaseSync } from './sqlite';
import { normalizeText } from './reducer';
import { normalizeAnimeIdentityKey } from './storage';
import { nowMs } from './time';
import { toDbTimestamp } from './query-shared';
import type { JellyfinLinkRepairSummary } from './types';

type LegacyJellyfinVideoRow = {
  video_id: number;
  video_key: string;
  anime_id: number | null;
  anime_assignment_locked: number;
  source_url: string | null;
  canonical_title: string;
};

type JellyfinTargetVideoRow = {
  video_id: number;
  anime_id: number | null;
  anime_assignment_locked: number;
  canonical_title: string;
  parsed_basename: string | null;
  parsed_title: string | null;
  parsed_season: number | null;
  parsed_episode: number | null;
  parser_source: string | null;
  parser_confidence: number | null;
  parse_metadata_json: string | null;
};

type LeakedAnimeTitleRow = {
  anime_id: number;
  canonical_title: string;
  normalized_title_key: string;
  title_romaji: string | null;
  title_english: string | null;
  title_native: string | null;
  linked_video_title: string | null;
};

function looksLikeLeakedJellyfinTitle(value: string | null): boolean {
  if (!value) return false;
  const lowered = value.toLowerCase();
  const hasApiKey = /api[\s_-]*key(?:\s|=|$)/i.test(value);
  return (
    hasApiKey &&
    (lowered.includes('stream?') ||
      lowered.includes('/stream?') ||
      lowered.includes('/videos/') ||
      lowered.includes('mediasourceid'))
  );
}

function chooseSafeAnimeTitle(row: LeakedAnimeTitleRow): string | null {
  const candidates = [
    row.title_english,
    row.title_romaji,
    row.title_native,
    row.linked_video_title?.replace(/^\[Jellyfin\/direct]\s*/i, ''),
  ];
  for (const candidate of candidates) {
    const normalized = candidate?.trim();
    if (normalized && !looksLikeLeakedJellyfinTitle(normalized)) {
      return normalized;
    }
  }
  return null;
}

function parseLegacyJellyfinStreamUrl(value: string | null): URL | null {
  if (!value) return null;
  const trimmed = value.trim();
  const urlText = trimmed.startsWith('remote:') ? trimmed.slice('remote:'.length) : trimmed;
  try {
    const url = new URL(urlText);
    const pathSegments = url.pathname.split('/').filter(Boolean);
    const videosIndex = pathSegments.findIndex((segment) => segment.toLowerCase() === 'videos');
    if (
      videosIndex < 0 ||
      pathSegments[videosIndex + 1] === undefined ||
      pathSegments[videosIndex + 2]?.toLowerCase() !== 'stream'
    ) {
      return null;
    }
    if (!url.searchParams.has('api_key')) {
      return null;
    }
    return url;
  } catch {
    return null;
  }
}

function buildJellyfinStatsUrlFromLegacyStream(url: URL): string | null {
  const pathSegments = url.pathname.split('/').filter(Boolean);
  const videosIndex = pathSegments.findIndex((segment) => segment.toLowerCase() === 'videos');
  const itemId = normalizeText(pathSegments[videosIndex + 1]);
  if (!itemId) return null;
  return `jellyfin://${url.host}/item/${encodeURIComponent(itemId)}`;
}

function buildSanitizedJellyfinVideoKey(
  db: DatabaseSync,
  videoId: number,
  statsUrl: string,
): string {
  const baseKey = `remote:${statsUrl}`;
  const existing = db
    .prepare('SELECT video_id FROM imm_videos WHERE video_key = ?')
    .get(baseKey) as { video_id: number } | null;
  if (!existing || existing.video_id === videoId) {
    return baseKey;
  }
  return `${baseKey}#legacy-${videoId}`;
}

function repairLeakedJellyfinAnimeTitles(db: DatabaseSync, currentTimestamp: string): number {
  const candidates = (
    db
      .prepare(
        `
          SELECT
            a.anime_id,
            a.normalized_title_key,
            a.canonical_title,
            a.title_romaji,
            a.title_english,
            a.title_native,
            (
              SELECT v.canonical_title
              FROM imm_videos v
              WHERE v.anime_id = a.anime_id
                AND v.canonical_title NOT LIKE '%api_key=%'
                AND lower(v.canonical_title) NOT LIKE '%api key%'
              ORDER BY v.LAST_UPDATE_DATE DESC, v.video_id DESC
              LIMIT 1
            ) AS linked_video_title
          FROM imm_anime a
          WHERE a.canonical_title LIKE '%api_key=%'
             OR lower(a.canonical_title) LIKE '%api key%'
             OR lower(a.normalized_title_key) LIKE '%api key%'
        `,
      )
      .all() as LeakedAnimeTitleRow[]
  ).filter(
    (row) =>
      looksLikeLeakedJellyfinTitle(row.canonical_title) ||
      looksLikeLeakedJellyfinTitle(row.normalized_title_key),
  );

  let repaired = 0;
  for (const candidate of candidates) {
    const replacementTitle = chooseSafeAnimeTitle(candidate);
    if (!replacementTitle) {
      continue;
    }
    const replacementKey = normalizeAnimeIdentityKey(replacementTitle);
    if (!replacementKey) {
      continue;
    }
    const existing = db
      .prepare(
        `
          SELECT anime_id
          FROM imm_anime
          WHERE normalized_title_key = ?
            AND anime_id != ?
        `,
      )
      .get(replacementKey, candidate.anime_id) as { anime_id: number } | null;
    if (existing) {
      const videoUpdate = db
        .prepare(
          `
            UPDATE imm_videos
            SET anime_id = ?, LAST_UPDATE_DATE = ?
            WHERE anime_id = ?
          `,
        )
        .run(existing.anime_id, currentTimestamp, candidate.anime_id) as { changes: number };
      const subtitleUpdate = db
        .prepare(
          `
            UPDATE imm_subtitle_lines
            SET anime_id = ?, LAST_UPDATE_DATE = ?
            WHERE anime_id = ?
          `,
        )
        .run(existing.anime_id, currentTimestamp, candidate.anime_id) as { changes: number };
      const animeDelete = db
        .prepare(
          `
            DELETE FROM imm_anime
            WHERE anime_id = ?
              AND NOT EXISTS (SELECT 1 FROM imm_videos WHERE anime_id = ?)
              AND NOT EXISTS (SELECT 1 FROM imm_subtitle_lines WHERE anime_id = ?)
          `,
        )
        .run(candidate.anime_id, candidate.anime_id, candidate.anime_id) as { changes: number };
      if (videoUpdate.changes > 0 || subtitleUpdate.changes > 0) {
        repaired += 1;
      } else if (animeDelete.changes > 0) {
        repaired += 1;
      }
      continue;
    }
    const updated = db
      .prepare(
        `
          UPDATE imm_anime
          SET
            normalized_title_key = ?,
            canonical_title = ?,
            LAST_UPDATE_DATE = ?
          WHERE anime_id = ?
        `,
      )
      .run(replacementKey, replacementTitle, currentTimestamp, candidate.anime_id) as {
      changes: number;
    };
    if (updated.changes > 0) {
      repaired += 1;
    }
  }
  return repaired;
}

function repairLeakedJellyfinVideoParseMetadata(
  db: DatabaseSync,
  currentTimestamp: string,
): number {
  const updated = db
    .prepare(
      `
        UPDATE imm_videos
        SET
          parsed_basename = NULL,
          parsed_title = NULL,
          parse_metadata_json = NULL,
          parser_source = CASE
            WHEN parser_source = 'guessit' THEN 'jellyfin'
            ELSE parser_source
          END,
          LAST_UPDATE_DATE = ?
        WHERE source_type = 2
          AND (
            parsed_basename LIKE '%api_key=%'
            OR lower(parsed_basename) LIKE '%api key%'
            OR parsed_title LIKE '%api_key=%'
            OR lower(parsed_title) LIKE '%api key%'
            OR parse_metadata_json LIKE '%api_key=%'
            OR lower(parse_metadata_json) LIKE '%api key%'
          )
      `,
    )
    .run(currentTimestamp) as { changes: number };
  return updated.changes;
}

export function repairJellyfinStreamVideoLinks(db: DatabaseSync): JellyfinLinkRepairSummary {
  const candidates = db
    .prepare(
      `
        SELECT
          video_id,
          video_key,
          anime_id,
          anime_assignment_locked,
          source_url,
          canonical_title
        FROM imm_videos
        WHERE source_type = 2
          AND (
            video_key LIKE '%api_key=%'
            OR lower(video_key) LIKE '%api key%'
            OR source_url LIKE '%api_key=%'
            OR lower(source_url) LIKE '%api key%'
            OR canonical_title LIKE '%api_key=%'
            OR lower(canonical_title) LIKE '%api key%'
          )
      `,
    )
    .all() as LegacyJellyfinVideoRow[];

  const summary: JellyfinLinkRepairSummary = {
    scanned: candidates.length,
    repaired: 0,
  };
  if (candidates.length === 0) {
    const currentTimestamp = toDbTimestamp(nowMs());
    const repaired =
      repairLeakedJellyfinAnimeTitles(db, currentTimestamp) +
      repairLeakedJellyfinVideoParseMetadata(db, currentTimestamp);
    summary.repaired += repaired;
    return summary;
  }

  const currentTimestamp = toDbTimestamp(nowMs());
  db.exec('BEGIN IMMEDIATE');
  try {
    for (const candidate of candidates) {
      const legacyUrl =
        parseLegacyJellyfinStreamUrl(candidate.source_url) ??
        parseLegacyJellyfinStreamUrl(candidate.video_key);
      if (!legacyUrl) {
        continue;
      }
      const statsUrl = buildJellyfinStatsUrlFromLegacyStream(legacyUrl);
      if (!statsUrl) {
        continue;
      }
      const sanitizedVideoKey = buildSanitizedJellyfinVideoKey(db, candidate.video_id, statsUrl);
      const sanitizedCanonicalTitle = looksLikeLeakedJellyfinTitle(candidate.canonical_title)
        ? 'Jellyfin Video'
        : candidate.canonical_title;
      const target = db
        .prepare(
          `
            SELECT
              video_id,
              anime_id,
              anime_assignment_locked,
              canonical_title,
              parsed_basename,
              parsed_title,
              parsed_season,
              parsed_episode,
              parser_source,
              parser_confidence,
              parse_metadata_json
            FROM imm_videos
            WHERE video_id != ?
              AND (video_key = ? OR source_url = ?)
            ORDER BY parser_source = 'jellyfin' DESC, video_id DESC
            LIMIT 1
          `,
        )
        .get(candidate.video_id, `remote:${statsUrl}`, statsUrl) as JellyfinTargetVideoRow | null;
      if (!target) {
        const updated = db
          .prepare(
            `
              UPDATE imm_videos
              SET
                video_key = ?,
                source_url = ?,
                canonical_title = ?,
                parser_source = COALESCE(parser_source, 'jellyfin'),
                LAST_UPDATE_DATE = ?
              WHERE video_id = ?
                AND (video_key != ? OR source_url != ? OR canonical_title != ?)
            `,
          )
          .run(
            sanitizedVideoKey,
            statsUrl,
            sanitizedCanonicalTitle,
            currentTimestamp,
            candidate.video_id,
            sanitizedVideoKey,
            statsUrl,
            sanitizedCanonicalTitle,
          ) as { changes: number };
        if (updated.changes > 0) {
          summary.repaired += 1;
        }
        continue;
      }

      const assignmentAnimeId =
        candidate.anime_assignment_locked === 1 ? candidate.anime_id : target.anime_id;
      // A lock without an assignment is meaningless and would pin later
      // relinking to nothing, so never carry the flag onto a NULL anime_id.
      const assignmentLocked =
        assignmentAnimeId !== null &&
        (candidate.anime_assignment_locked === 1 || target.anime_assignment_locked === 1)
          ? 1
          : 0;
      db.prepare(
        `
          UPDATE imm_videos
          SET
            video_key = ?,
            anime_id = ?,
            anime_assignment_locked = ?,
            canonical_title = ?,
            source_url = ?,
            parsed_basename = ?,
            parsed_title = ?,
            parsed_season = ?,
            parsed_episode = ?,
            parser_source = ?,
            parser_confidence = ?,
            parse_metadata_json = ?,
            LAST_UPDATE_DATE = ?
          WHERE video_id = ?
        `,
      ).run(
        sanitizedVideoKey,
        assignmentAnimeId,
        assignmentLocked,
        target.canonical_title,
        statsUrl,
        target.parsed_basename,
        target.parsed_title,
        target.parsed_season,
        target.parsed_episode,
        target.parser_source,
        target.parser_confidence,
        target.parse_metadata_json,
        currentTimestamp,
        candidate.video_id,
      );
      db.prepare(
        `
          UPDATE imm_subtitle_lines
          SET anime_id = ?, LAST_UPDATE_DATE = ?
          WHERE video_id = ?
        `,
      ).run(assignmentAnimeId, currentTimestamp, candidate.video_id);
      summary.repaired += 1;
    }
    summary.repaired += repairLeakedJellyfinAnimeTitles(db, currentTimestamp);
    summary.repaired += repairLeakedJellyfinVideoParseMetadata(db, currentTimestamp);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }

  return summary;
}
