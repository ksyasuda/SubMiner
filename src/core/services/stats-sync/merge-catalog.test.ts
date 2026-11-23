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

test('franchise override follows whichever side set or cleared it last', () => {
  const adapt = (db: DatabaseSync) => ({
    query: (sql: string) => db.prepare(sql),
    exec: (sql: string) => {
      db.exec(sql);
    },
    close: () => {
      db.close();
    },
  });
  const cases: Array<{
    name: string;
    local: [number | null, string | null];
    remote: [number | null, string | null];
    expected: number | null;
  }> = [
    {
      name: 'newer remote join wins',
      local: [null, null],
      remote: [20832, '2000'],
      expected: 20832,
    },
    { name: 'newer remote clear wins', local: [0, '1000'], remote: [null, '2000'], expected: null },
    { name: 'older remote loses', local: [0, '3000'], remote: [20832, '2000'], expected: 0 },
    { name: 'untouched remote keeps local', local: [0, '1000'], remote: [null, null], expected: 0 },
  ];
  for (const c of cases) {
    const local = new Database(':memory:');
    const remote = new Database(':memory:');
    try {
      for (const [db, [override, updatedAt]] of [
        [local, c.local],
        [remote, c.remote],
      ] as const) {
        ensureSchema(db);
        db.prepare(
          `INSERT INTO imm_anime(normalized_title_key, canonical_title, anilist_id, group_override, group_override_updated_at)
          VALUES ('overlord ii', 'Overlord II', 98437, ?, ?)`,
        ).run(override, updatedAt);
      }
      mergeAnime(adapt(local), adapt(remote), createEmptyMergeSummary());
      const row = local.prepare('SELECT group_override AS value FROM imm_anime').get() as {
        value: number | null;
      };
      assert.equal(row.value, c.expected, c.name);
    } finally {
      local.close();
      remote.close();
    }
  }

  const local = new Database(':memory:');
  const remote = new Database(':memory:');
  try {
    ensureSchema(local);
    ensureSchema(remote);
    remote
      .prepare(
        `INSERT INTO imm_anime(normalized_title_key, canonical_title, group_override, group_override_updated_at)
        VALUES ('specials', 'Specials', 20832, '1000')`,
      )
      .run();
    mergeAnime(adapt(local), adapt(remote), createEmptyMergeSummary());
    const row = local
      .prepare(
        'SELECT group_override AS value, group_override_updated_at AS updatedAt FROM imm_anime',
      )
      .get() as { value: number | null; updatedAt: string | null };
    assert.deepEqual([row.value, row.updatedAt], [20832, '1000']);
  } finally {
    local.close();
    remote.close();
  }
});
