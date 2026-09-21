import test from 'node:test';
import assert from 'node:assert/strict';
import { Database, type DatabaseSync } from '../immersion-tracker/sqlite';
import { ensureSchema } from '../immersion-tracker/storage';
import { mergeAnime } from './merge-catalog';
import { createEmptyMergeSummary } from './shared';

const identities = {
  unlinked: [null, null, null],
  anilist: [42, null, null],
  tmdb: [null, 12, 'tv'],
  otherTmdb: [null, 13, 'tv'],
} as const;

for (const [localKind, remoteKind, sameEntry] of [
  ['anilist', 'tmdb', false],
  ['tmdb', 'anilist', false],
  ['tmdb', 'otherTmdb', false],
  ['unlinked', 'tmdb', true],
  ['unlinked', 'anilist', true],
  ['tmdb', 'unlinked', true],
  ['tmdb', 'tmdb', true],
  ['anilist', 'anilist', true],
] as const) {
  test(`catalog title match: ${localKind} with ${remoteKind}`, () => {
    const local = new Database(':memory:');
    const remote = new Database(':memory:');
    const adapt = (db: DatabaseSync) => ({
      query: (sql: string) => db.prepare(sql),
      exec: (sql: string) => {
        db.exec(sql);
      },
      close: () => {
        db.close();
      },
    });
    try {
      for (const [db, kind] of [
        [local, localKind],
        [remote, remoteKind],
      ] as const) {
        ensureSchema(db);
        db.prepare(
          `INSERT INTO imm_anime(normalized_title_key, canonical_title, anilist_id, tmdb_id, tmdb_type, CREATED_DATE, LAST_UPDATE_DATE)
          VALUES ('same title', 'Same title', ?, ?, ?, 1000, 1000)`,
        ).run(...identities[kind]);
      }
      const summary = createEmptyMergeSummary();
      const map = mergeAnime(adapt(local), adapt(remote), summary);
      assert.equal(map.get(1) === 1, sameEntry);
      assert.equal(summary.animeAdded, sameEntry ? 0 : 1);
      const again = createEmptyMergeSummary();
      assert.equal(mergeAnime(adapt(local), adapt(remote), again).get(1), map.get(1));
      assert.equal(again.animeAdded, 0);
    } finally {
      local.close();
      remote.close();
    }
  });
}
