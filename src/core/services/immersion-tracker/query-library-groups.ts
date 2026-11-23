import type {
  AnimeGroupOverrideMode,
  AnimeGroupPlacement,
  AnimeLibraryGroup,
  AnimeLibraryGroupPart,
} from '../../../types/stats-wire';
import {
  anilistNodeTitle,
  buildFranchiseGraph,
  type FranchiseGraph,
} from '../anilist/franchise-graph';
import { readAnilistRelationCache } from './anilist-relation-cache';
import { toDbTimestamp } from './query-shared';
import { getAnimeLibrary } from './query-library';
import type { DatabaseSync } from './sqlite';
import { nowMs } from './time';
import type { AnimeLibraryRow } from './types';

/** `imm_anime.group_override` value that keeps an entry on its own. */
export const GROUP_OVERRIDE_STANDALONE = 0;

export interface GroupableLibraryRow extends AnimeLibraryRow {
  groupOverride: number | null;
}

function overrideMode(row: GroupableLibraryRow): AnimeGroupOverrideMode {
  if (row.mediaKind !== 'anime' || row.groupOverride === null || row.groupOverride < 0) {
    return 'anilist';
  }
  return row.groupOverride === GROUP_OVERRIDE_STANDALONE ? 'standalone' : 'joined';
}

/**
 * The card an entry belongs to: positive keys are AniList ids (the franchise
 * root once its relations are cached, the entry's own id until then), negative
 * keys are `-animeId` for entries listed on their own. Overrides only apply to
 * anime entries; live-action and YouTube rows always stand alone.
 */
function groupKeyFor(row: GroupableLibraryRow, graph: FranchiseGraph): number {
  const mode = overrideMode(row);
  if (mode === 'standalone') return -row.animeId;
  if (mode === 'joined') return graph.rootOf(row.groupOverride!) ?? row.groupOverride!;
  if (row.mediaKind === 'anime' && row.anilistId !== null) {
    return graph.rootOf(row.anilistId) ?? row.anilistId;
  }
  return -row.animeId;
}

function toPart(row: GroupableLibraryRow, graph: FranchiseGraph): AnimeLibraryGroupPart {
  const { groupOverride: _groupOverride, ...item } = row;
  const node =
    row.mediaKind === 'anime' && row.anilistId !== null ? graph.node(row.anilistId) : undefined;
  return { ...item, format: node?.format ?? null, startDate: node?.startDate ?? null };
}

// Release order; entries AniList has no date for go last, in library order.
function compareParts(a: AnimeLibraryGroupPart, b: AnimeLibraryGroupPart): number {
  return (
    (a.startDate ?? Number.MAX_SAFE_INTEGER) - (b.startDate ?? Number.MAX_SAFE_INTEGER) ||
    a.animeId - b.animeId
  );
}

/**
 * Groups library rows into franchise cards. A card with one part keeps that
 * entry's own title, so ungrouped entries render exactly as before; a card with
 * several parts is named after the franchise root, even when the root itself
 * is not in the library.
 */
export function groupAnimeLibrary(
  rows: GroupableLibraryRow[],
  graph: FranchiseGraph,
): AnimeLibraryGroup[] {
  const byKey = new Map<number, GroupableLibraryRow[]>();
  for (const row of rows) {
    const key = groupKeyFor(row, graph);
    const members = byKey.get(key);
    if (members) members.push(row);
    else byKey.set(key, [row]);
  }

  const groups: AnimeLibraryGroup[] = [];
  for (const [groupKey, members] of byKey) {
    const parts = members.map((row) => toPart(row, graph)).sort(compareParts);
    const first = parts[0]!;
    const rootNode = groupKey > 0 ? graph.node(groupKey) : undefined;
    const title =
      parts.length === 1
        ? first.canonicalTitle
        : ((rootNode ? anilistNodeTitle(rootNode) : null) ??
          parts.find((part) => part.anilistId === groupKey)?.canonicalTitle ??
          first.canonicalTitle);
    groups.push({
      groupKey,
      title,
      mediaKind: first.mediaKind,
      coverAnimeId: first.animeId,
      totalSessions: parts.reduce((sum, part) => sum + part.totalSessions, 0),
      totalActiveMs: parts.reduce((sum, part) => sum + part.totalActiveMs, 0),
      totalCards: parts.reduce((sum, part) => sum + part.totalCards, 0),
      totalTokensSeen: parts.reduce((sum, part) => sum + part.totalTokensSeen, 0),
      episodeCount: parts.reduce((sum, part) => sum + part.episodeCount, 0),
      lastWatchedMs: Math.max(...parts.map((part) => part.lastWatchedMs)),
      parts,
    });
  }
  return groups.sort((a, b) => b.totalActiveMs - a.totalActiveMs || a.groupKey - b.groupKey);
}

function readGroupableLibrary(db: DatabaseSync): GroupableLibraryRow[] {
  const overrides = new Map(
    (
      db
        .prepare(
          'SELECT anime_id AS animeId, group_override AS groupOverride FROM imm_anime WHERE group_override IS NOT NULL',
        )
        .all() as Array<{ animeId: number; groupOverride: number }>
    ).map((row) => [row.animeId, row.groupOverride] as const),
  );
  return getAnimeLibrary(db).map((row) => ({
    ...row,
    groupOverride: overrides.get(row.animeId) ?? null,
  }));
}

export function getAnimeLibraryGroups(db: DatabaseSync): AnimeLibraryGroup[] {
  return groupAnimeLibrary(
    readGroupableLibrary(db),
    buildFranchiseGraph(readAnilistRelationCache(db).values()),
  );
}

/** The card an entry is listed under, or null when the entry is not in the library list. */
export function getAnimeGroupPlacement(
  db: DatabaseSync,
  animeId: number,
): AnimeGroupPlacement | null {
  const rows = readGroupableLibrary(db);
  const row = rows.find((candidate) => candidate.animeId === animeId);
  if (!row) return null;
  const group = groupAnimeLibrary(
    rows,
    buildFranchiseGraph(readAnilistRelationCache(db).values()),
  ).find((candidate) => candidate.parts.some((part) => part.animeId === animeId));
  if (!group) return null;
  return {
    groupKey: group.groupKey,
    title: group.title,
    partCount: group.parts.length,
    overrideMode: overrideMode(row),
  };
}

/** Where to put an entry: with another entry's franchise, on its own, or where AniList has it. */
export type AnimeGroupOverrideTarget =
  | { mode: 'join'; targetAnimeId: number }
  | { mode: 'standalone' }
  | { mode: 'anilist' };

export type AnimeGroupOverrideResult =
  | { ok: true }
  | { ok: false; reason: 'not-found' | 'unsupported-kind' | 'target-not-grouped' };

/**
 * Stores the user's grouping choice for one anime entry. Joining stores the
 * target's current card key, which is always an AniList id: entries without
 * one cannot be joined (merge them instead).
 */
export function setAnimeGroupOverride(
  db: DatabaseSync,
  animeId: number,
  target: AnimeGroupOverrideTarget,
): AnimeGroupOverrideResult {
  const row = db
    .prepare('SELECT media_kind AS mediaKind FROM imm_anime WHERE anime_id = ?')
    .get(animeId) as { mediaKind: string } | undefined;
  if (!row) return { ok: false, reason: 'not-found' };
  if (row.mediaKind !== 'anime') return { ok: false, reason: 'unsupported-kind' };

  let value: number | null = null;
  if (target.mode === 'standalone') value = GROUP_OVERRIDE_STANDALONE;
  if (target.mode === 'join') {
    const placement = getAnimeGroupPlacement(db, target.targetAnimeId);
    if (!placement) return { ok: false, reason: 'not-found' };
    if (placement.groupKey <= 0) return { ok: false, reason: 'target-not-grouped' };
    value = placement.groupKey;
  }
  db.prepare(
    'UPDATE imm_anime SET group_override = ?, group_override_updated_at = ? WHERE anime_id = ?',
  ).run(value, toDbTimestamp(nowMs()), animeId);
  return { ok: true };
}
