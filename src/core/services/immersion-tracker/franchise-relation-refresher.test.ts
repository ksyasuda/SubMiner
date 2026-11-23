import assert from 'node:assert/strict';
import test from 'node:test';
import type { AnilistMediaNode } from '../anilist/franchise-graph';
import { readAnilistRelationCache } from './anilist-relation-cache';
import {
  createFranchiseRelationRefresher,
  FRANCHISE_RELATION_STALE_MS,
} from './franchise-relation-refresher';
import { Database } from './sqlite';
import { ensureSchema } from './storage';

const silent = { info: () => {}, warn: () => {} };

function media(anilistId: number, relations: Array<[string, number]>, fetchedAtMs: number) {
  return {
    anilistId,
    format: 'TV',
    titleRomaji: null,
    titleEnglish: null,
    titleNative: null,
    startDate: null,
    episodes: null,
    relations: relations.map(([relationType, targetId]) => ({ relationType, targetId })),
    fetchedAtMs,
  } satisfies AnilistMediaNode;
}

test('refresh walks from linked entries and override targets, then keeps the cache on failure', async () => {
  const db = new Database(':memory:');
  try {
    ensureSchema(db);
    db.exec(`INSERT INTO imm_anime(normalized_title_key, canonical_title, anilist_id, media_kind)
      VALUES ('a', 'A', 1, 'anime'), ('c', 'C', 77, 'youtube');
      INSERT INTO imm_anime(normalized_title_key, canonical_title, group_override)
      VALUES ('b', 'B', 50)`);
    let clock = 1_000;
    let failing = false;
    const asked: number[][] = [];
    const refresher = createFranchiseRelationRefresher({
      db,
      logger: silent,
      now: () => clock,
      fetcher: {
        fetchMedia: async (ids) => {
          if (failing) throw new Error('AniList down');
          asked.push([...ids].sort((a, b) => a - b));
          return ids.map((id) => media(id, id === 1 ? [['SEQUEL', 2]] : [], clock));
        },
      },
    });

    refresher.requestRefresh();
    await refresher.whenIdle();
    assert.deepEqual(asked, [[1, 50], [2]]);
    assert.deepEqual(
      [...readAnilistRelationCache(db).keys()].sort((a, b) => a - b),
      [1, 2, 50],
    );

    // Throttled until forced; everything is fresh, so a forced run asks nothing.
    refresher.requestRefresh({ force: true });
    await refresher.whenIdle();
    assert.equal(asked.length, 2);

    clock += FRANCHISE_RELATION_STALE_MS;
    failing = true;
    refresher.requestRefresh({ force: true });
    await refresher.whenIdle();
    assert.deepEqual(readAnilistRelationCache(db).get(1)?.relations, [
      { relationType: 'SEQUEL', targetId: 2 },
    ]);
  } finally {
    db.close();
  }
});
