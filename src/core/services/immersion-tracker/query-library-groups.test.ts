import assert from 'node:assert/strict';
import test from 'node:test';
import { buildFranchiseGraph, type AnilistMediaNode } from '../anilist/franchise-graph';
import { readAnilistRelationCache, writeAnilistRelationNodes } from './anilist-relation-cache';
import {
  getAnimeGroupPlacement,
  getAnimeLibraryGroups,
  groupAnimeLibrary,
  setAnimeGroupOverride,
  type GroupableLibraryRow,
} from './query-library-groups';
import { Database, type DatabaseSync } from './sqlite';
import { ensureSchema } from './storage';
import { SCHEMA_VERSION } from './types';

function node(
  anilistId: number,
  format: string,
  startDate: number,
  relations: Array<[string, number]>,
  titleEnglish: string | null = null,
): AnilistMediaNode {
  return {
    anilistId,
    format,
    titleRomaji: `romaji ${anilistId}`,
    titleEnglish,
    titleNative: null,
    startDate,
    episodes: null,
    relations: relations.map(([relationType, targetId]) => ({ relationType, targetId })),
    fetchedAtMs: 1,
  };
}

// Owned: season 2 and a film. The first season (the root) is not in the library.
const overlordNodes = [
  node(
    20832,
    'TV',
    20150707,
    [
      ['SEQUEL', 98437],
      ['SUMMARY', 98873],
    ],
    'Overlord',
  ),
  node(98873, 'MOVIE', 20170225, [['PARENT', 20832]]),
  node(98437, 'TV', 20180109, [['PREQUEL', 20832]]),
];

function row(
  animeId: number,
  anilistId: number | null,
  overrides: Partial<GroupableLibraryRow> = {},
): GroupableLibraryRow {
  return {
    mediaKind: 'anime',
    animeId,
    canonicalTitle: `Entry ${animeId}`,
    anilistId,
    tmdbId: null,
    tmdbType: null,
    totalSessions: 1,
    totalActiveMs: 1000 * animeId,
    totalCards: animeId,
    totalTokensSeen: 10,
    episodeCount: 2,
    episodesTotal: null,
    lastWatchedMs: 100 * animeId,
    groupOverride: null,
    ...overrides,
  };
}

test('owned parts join under the unowned root, in release order, with summed totals', () => {
  const graph = buildFranchiseGraph(overlordNodes);
  const groups = groupAnimeLibrary([row(1, 98437), row(2, 98873), row(3, null)], graph);
  assert.equal(groups.length, 2);
  const overlord = groups.find((group) => group.groupKey === 20832)!;
  assert.equal(overlord.title, 'Overlord');
  assert.deepEqual(
    overlord.parts.map((part) => [part.animeId, part.format, part.startDate]),
    [
      [2, 'MOVIE', 20170225],
      [1, 'TV', 20180109],
    ],
  );
  assert.equal(overlord.coverAnimeId, 2);
  assert.deepEqual(
    [overlord.totalActiveMs, overlord.totalCards, overlord.episodeCount, overlord.lastWatchedMs],
    [3000, 3, 4, 200],
  );
});

test('entries without AniList data stay single cards that look like the entry', () => {
  const graph = buildFranchiseGraph([]);
  const lone = row(7, 555, { canonicalTitle: 'Bocchi the Rock!' });
  const [linked, unlinked, channel] = groupAnimeLibrary(
    [lone, row(3, null), row(2, null, { mediaKind: 'youtube', groupOverride: 555 })],
    graph,
  );
  assert.deepEqual(
    [linked!.groupKey, linked!.title, linked!.coverAnimeId, linked!.parts.length],
    [555, 'Bocchi the Rock!', 7, 1],
  );
  const { groupOverride: _override, ...item } = lone;
  assert.deepEqual(linked!.parts[0], { ...item, format: null, startDate: null });
  assert.equal(unlinked!.groupKey, -3);
  // Overrides never apply to YouTube or live-action entries.
  assert.equal(channel!.groupKey, -2);
});

test('overrides move an entry to another franchise or keep it on its own', () => {
  const graph = buildFranchiseGraph(overlordNodes);
  const groups = groupAnimeLibrary(
    [
      row(1, 98437),
      row(2, 98873, { groupOverride: 0 }),
      // Unlinked specials moved in by the user, pointing at any part of the franchise.
      row(3, null, { groupOverride: 98437 }),
    ],
    graph,
  );
  assert.deepEqual(
    groups
      .map((group) => [group.groupKey, group.parts.map((part) => part.animeId)] as const)
      .sort((a, b) => a[0] - b[0]),
    [
      [-2, [2]],
      [20832, [1, 3]],
    ],
  );
});

function seed(db: DatabaseSync, title: string, anilistId: number | null): number {
  const animeId = Number(
    db
      .prepare(
        `INSERT INTO imm_anime(normalized_title_key, canonical_title, anilist_id) VALUES (?, ?, ?)`,
      )
      .run(title.toLowerCase(), title, anilistId).lastInsertRowid,
  );
  db.prepare(
    `INSERT INTO imm_videos(video_key, anime_id, canonical_title, source_type, duration_ms)
     VALUES (?, ?, ?, 1, 0)`,
  ).run(`key:${title}`, animeId, title);
  db.prepare(
    'INSERT INTO imm_lifetime_anime(anime_id, total_active_ms, total_cards) VALUES (?, 60000, 1)',
  ).run(animeId);
  return animeId;
}

test('library groups and overrides round-trip through the database', () => {
  const db = new Database(':memory:');
  try {
    ensureSchema(db);
    const season2 = seed(db, 'Overlord II', 98437);
    const film = seed(db, 'The Undead King', 98873);
    const specials = seed(db, 'Overlord specials', null);
    const other = seed(db, 'Bocchi the Rock!', 130003);

    // No cached relations yet: every entry is its own card.
    assert.equal(getAnimeLibraryGroups(db).length, 4);

    writeAnilistRelationNodes(db, overlordNodes);
    assert.deepEqual(readAnilistRelationCache(db).get(20832), overlordNodes[0]);
    assert.deepEqual(getAnimeGroupPlacement(db, film), {
      groupKey: 20832,
      title: 'Overlord',
      partCount: 2,
      overrideMode: 'anilist',
    });

    assert.deepEqual(
      setAnimeGroupOverride(db, specials, { mode: 'join', targetAnimeId: season2 }),
      {
        ok: true,
      },
    );
    assert.deepEqual(setAnimeGroupOverride(db, film, { mode: 'standalone' }), { ok: true });
    assert.deepEqual(setAnimeGroupOverride(db, other, { mode: 'join', targetAnimeId: film }), {
      ok: false,
      reason: 'target-not-grouped',
    });
    assert.deepEqual(setAnimeGroupOverride(db, 999, { mode: 'standalone' }), {
      ok: false,
      reason: 'not-found',
    });
    const stored = db
      .prepare('SELECT group_override AS value FROM imm_anime WHERE anime_id = ?')
      .get(specials) as { value: number };
    assert.equal(stored.value, 20832);
    assert.deepEqual(getAnimeGroupPlacement(db, specials), {
      groupKey: 20832,
      title: 'Overlord',
      partCount: 2,
      overrideMode: 'joined',
    });
    assert.equal(getAnimeGroupPlacement(db, film)?.overrideMode, 'standalone');

    setAnimeGroupOverride(db, film, { mode: 'anilist' });
    assert.equal(getAnimeGroupPlacement(db, film)?.partCount, 3);
  } finally {
    db.close();
  }
});

test('schema 27 upgrade adds the relation cache and override columns', () => {
  const db = new Database(':memory:');
  try {
    ensureSchema(db);
    db.exec(`DROP TABLE imm_anilist_relations; DROP TABLE imm_anilist_media;
      ALTER TABLE imm_anime DROP COLUMN group_override;
      ALTER TABLE imm_anime DROP COLUMN group_override_updated_at;`);
    db.prepare('UPDATE imm_schema_version SET schema_version = ?').run(SCHEMA_VERSION - 1);
    ensureSchema(db);
    const columns = (db.prepare('PRAGMA table_info(imm_anime)').all() as Array<{ name: string }>)
      .map((column) => column.name)
      .filter((name) => name.startsWith('group_override'));
    assert.deepEqual(columns, ['group_override', 'group_override_updated_at']);
    assert.equal(readAnilistRelationCache(db).size, 0);
  } finally {
    db.close();
  }
});
