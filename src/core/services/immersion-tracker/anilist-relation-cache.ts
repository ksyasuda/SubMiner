import type { AnilistMediaNode } from '../anilist/franchise-graph';
import type { DatabaseSync } from './sqlite';
import { fromDbTimestamp, toDbTimestamp } from './query-shared';

/**
 * Franchise grouping state. `imm_anilist_media` and `imm_anilist_relations`
 * cache AniList entries (owned or reached through relations) and every
 * anime-to-anime edge, whatever its type, so the grouping rule can change
 * without a refetch. The cache is rebuildable from AniList and is not synced.
 *
 * `imm_anime.group_override` is the user's choice for one library entry:
 * NULL follows AniList, 0 keeps the entry on its own, and a positive value is
 * an AniList id whose franchise the entry joins. `group_override_updated_at`
 * records when it was set (or cleared) so sync can keep the newer choice.
 */
export function ensureAnimeGroupingSchema(db: DatabaseSync): void {
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_anilist_media(
      anilist_id INTEGER PRIMARY KEY,
      format TEXT,
      title_romaji TEXT,
      title_english TEXT,
      title_native TEXT,
      start_date INTEGER,
      episodes INTEGER,
      fetched_at_ms TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS imm_anilist_relations(
      anilist_id INTEGER NOT NULL,
      relation_type TEXT NOT NULL,
      target_anilist_id INTEGER NOT NULL,
      PRIMARY KEY(anilist_id, relation_type, target_anilist_id),
      FOREIGN KEY(anilist_id) REFERENCES imm_anilist_media(anilist_id) ON DELETE CASCADE
    );
  `);
  const columns = new Set(
    (db.prepare('PRAGMA table_info(imm_anime)').all() as Array<{ name: string }>).map(
      (column) => column.name,
    ),
  );
  if (!columns.has('group_override')) {
    db.exec('ALTER TABLE imm_anime ADD COLUMN group_override INTEGER');
  }
  if (!columns.has('group_override_updated_at')) {
    db.exec('ALTER TABLE imm_anime ADD COLUMN group_override_updated_at TEXT');
  }
}

export function readAnilistRelationCache(db: DatabaseSync): Map<number, AnilistMediaNode> {
  const nodes = new Map<number, AnilistMediaNode>();
  const mediaRows = db
    .prepare(
      `SELECT anilist_id AS anilistId, format, title_romaji AS titleRomaji,
        title_english AS titleEnglish, title_native AS titleNative,
        start_date AS startDate, episodes, fetched_at_ms AS fetchedAtMs
      FROM imm_anilist_media`,
    )
    .all() as Array<Omit<AnilistMediaNode, 'relations' | 'fetchedAtMs'> & { fetchedAtMs: string }>;
  for (const row of mediaRows) {
    nodes.set(row.anilistId, {
      ...row,
      relations: [],
      fetchedAtMs: fromDbTimestamp(row.fetchedAtMs) ?? 0,
    });
  }
  const edgeRows = db
    .prepare(
      `SELECT anilist_id AS anilistId, relation_type AS relationType, target_anilist_id AS targetId
      FROM imm_anilist_relations
      ORDER BY anilist_id, relation_type, target_anilist_id`,
    )
    .all() as Array<{ anilistId: number; relationType: string; targetId: number }>;
  for (const edge of edgeRows) {
    nodes.get(edge.anilistId)?.relations.push({
      relationType: edge.relationType,
      targetId: edge.targetId,
    });
  }
  return nodes;
}

/** Replaces each node's cached row and edges in one transaction. */
export function writeAnilistRelationNodes(db: DatabaseSync, nodes: AnilistMediaNode[]): void {
  if (nodes.length === 0) return;
  const upsertMedia = db.prepare(`
    INSERT INTO imm_anilist_media(anilist_id, format, title_romaji, title_english, title_native,
      start_date, episodes, fetched_at_ms)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(anilist_id) DO UPDATE SET
      format = excluded.format,
      title_romaji = excluded.title_romaji,
      title_english = excluded.title_english,
      title_native = excluded.title_native,
      start_date = excluded.start_date,
      episodes = excluded.episodes,
      fetched_at_ms = excluded.fetched_at_ms
  `);
  const deleteEdges = db.prepare('DELETE FROM imm_anilist_relations WHERE anilist_id = ?');
  const insertEdge = db.prepare(`
    INSERT OR IGNORE INTO imm_anilist_relations(anilist_id, relation_type, target_anilist_id)
    VALUES (?, ?, ?)
  `);
  db.exec('BEGIN IMMEDIATE');
  try {
    for (const node of nodes) {
      upsertMedia.run(
        node.anilistId,
        node.format,
        node.titleRomaji,
        node.titleEnglish,
        node.titleNative,
        node.startDate,
        node.episodes,
        toDbTimestamp(node.fetchedAtMs),
      );
      deleteEdges.run(node.anilistId);
      for (const edge of node.relations) {
        insertEdge.run(node.anilistId, edge.relationType, edge.targetId);
      }
    }
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}
