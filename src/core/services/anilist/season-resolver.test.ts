import test from 'node:test';
import assert from 'node:assert/strict';

import {
  ANILIST_SEASON_RELATIONS_QUERY,
  ANILIST_SEASON_SEARCH_QUERY,
  pickByAirOrder,
  resolveAnilistSeasonMedia,
  stripSeasonSuffix,
  type AnilistSeasonMedia,
} from './season-resolver';

/** Real AniList payloads for "My Teen Romantic Comedy SNAFU", trimmed to used fields. */
const OREGAIRU_SEARCH: AnilistSeasonMedia[] = [
  {
    id: 14813,
    episodes: 13,
    format: 'TV',
    seasonYear: 2013,
    title: {
      romaji: 'Yahari Ore no Seishun Love Come wa Machigatteiru.',
      english: 'My Teen Romantic Comedy SNAFU',
      native: null,
    },
  },
  {
    id: 18753,
    episodes: 1,
    format: 'OVA',
    seasonYear: 2013,
    title: { romaji: null, english: 'My Teen Romantic Comedy SNAFU OVA', native: null },
  },
  {
    id: 108489,
    episodes: 12,
    format: 'TV',
    seasonYear: 2020,
    title: {
      romaji: 'Yahari Ore no Seishun Love Come wa Machigatteiru. Kan',
      english: 'My Teen Romantic Comedy SNAFU Climax!',
      native: null,
    },
  },
  {
    id: 20698,
    episodes: 13,
    format: 'TV',
    seasonYear: 2015,
    title: {
      romaji: 'Yahari Ore no Seishun Love Come wa Machigatteiru. Zoku',
      english: 'My Teen Romantic Comedy SNAFU TOO!',
      native: null,
    },
  },
];

const OREGAIRU_RELATIONS: Record<
  number,
  Array<{ relationType: string; node: AnilistSeasonMedia }>
> = {
  14813: [
    {
      relationType: 'SIDE_STORY',
      node: OREGAIRU_SEARCH[1]!,
    },
    {
      relationType: 'SEQUEL',
      node: OREGAIRU_SEARCH[3]!,
    },
  ],
  20698: [
    {
      relationType: 'PREQUEL',
      node: OREGAIRU_SEARCH[0]!,
    },
    {
      relationType: 'SEQUEL',
      node: OREGAIRU_SEARCH[2]!,
    },
  ],
  108489: [
    {
      relationType: 'PREQUEL',
      node: OREGAIRU_SEARCH[3]!,
    },
  ],
};

function createExecutor(
  search: AnilistSeasonMedia[],
  relations: Record<number, Array<{ relationType: string; node: AnilistSeasonMedia }>> = {},
) {
  const searches: string[] = [];
  const relationLookups: number[] = [];
  const execute = async <T>(query: string, variables: Record<string, unknown>): Promise<T> => {
    if (query === ANILIST_SEASON_SEARCH_QUERY) {
      searches.push(String(variables.search));
      return { Page: { media: search } } as T;
    }
    if (query === ANILIST_SEASON_RELATIONS_QUERY) {
      const id = Number(variables.id);
      relationLookups.push(id);
      return {
        Media: {
          relations: {
            edges: (relations[id] ?? []).map((edge) => ({
              relationType: edge.relationType,
              node: { ...edge.node, type: 'ANIME' },
            })),
          },
        },
      } as T;
    }
    throw new Error(`unexpected query: ${query}`);
  };
  return { execute, searches, relationLookups };
}

test('stripSeasonSuffix drops release-name season markers', () => {
  assert.equal(stripSeasonSuffix('Some Show Season 3'), 'Some Show');
  assert.equal(stripSeasonSuffix('Some Show S3'), 'Some Show');
  assert.equal(stripSeasonSuffix('Some Show 2nd Season'), 'Some Show');
  assert.equal(stripSeasonSuffix('Some Show'), 'Some Show');
});

test('resolves season 3 by walking sequel relations', async () => {
  const { execute, searches, relationLookups } = createExecutor(
    OREGAIRU_SEARCH,
    OREGAIRU_RELATIONS,
  );
  const result = await resolveAnilistSeasonMedia(
    { title: 'My Teen Romantic Comedy SNAFU', season: 3, episode: 1 },
    { execute },
  );

  assert.equal(result?.id, 108489);
  assert.equal(result?.title, 'My Teen Romantic Comedy SNAFU Climax!');
  assert.equal(result?.episodes, 12);
  assert.equal(result?.seasonResolved, true);
  assert.equal(result?.via, 'sequel-chain');
  // Searches the bare title only; "Season 3" returns nothing on AniList.
  assert.deepEqual(searches, ['My Teen Romantic Comedy SNAFU']);
  assert.deepEqual(relationLookups, [14813, 20698]);
});

test('season 1 resolves to the anchor without relation lookups', async () => {
  const { execute, relationLookups } = createExecutor(OREGAIRU_SEARCH, OREGAIRU_RELATIONS);
  const result = await resolveAnilistSeasonMedia(
    { title: 'My Teen Romantic Comedy SNAFU', season: 1, episode: 5 },
    { execute },
  );

  assert.equal(result?.id, 14813);
  assert.equal(result?.seasonResolved, true);
  assert.equal(result?.via, 'anchor');
  assert.deepEqual(relationLookups, []);
});

test('season 2 preserves the anchor exact-title evidence through a sequel resolution', async () => {
  const { execute } = createExecutor(OREGAIRU_SEARCH, OREGAIRU_RELATIONS);
  const result = await resolveAnilistSeasonMedia(
    { title: 'My Teen Romantic Comedy SNAFU', season: 2, episode: 1 },
    { execute },
  );

  assert.equal(result?.id, 20698);
  assert.equal(result?.via, 'sequel-chain');
  assert.equal(result?.exactTitleMatch, true);
});

test('reports an exact normalized synonym match as strong evidence', async () => {
  const { execute } = createExecutor([
    {
      id: 1,
      episodes: 12,
      format: 'TV',
      title: { english: 'Hitori Gotoh Story' },
      synonyms: ['BOCCHI THE ROCK'],
    },
  ]);

  const result = await resolveAnilistSeasonMedia({ title: 'Bocchi the Rock!' }, { execute });

  assert.equal(result?.exactTitleMatch, true);
});

test('reports a fuzzy-only search result as weak evidence', async () => {
  const { execute } = createExecutor([
    {
      id: 1,
      episodes: 12,
      format: 'TV',
      title: { english: 'Actual Show' },
    },
  ]);

  const result = await resolveAnilistSeasonMedia({ title: 'Unrelated Release' }, { execute });

  assert.equal(result?.exactTitleMatch, false);
});

test('strips a season marker already present in the parsed title', async () => {
  const { execute, searches } = createExecutor(OREGAIRU_SEARCH, OREGAIRU_RELATIONS);
  const result = await resolveAnilistSeasonMedia(
    { title: 'My Teen Romantic Comedy SNAFU Season 2', season: 2, episode: 3 },
    { execute },
  );

  assert.deepEqual(searches, ['My Teen Romantic Comedy SNAFU']);
  assert.equal(result?.id, 20698);
  assert.equal(result?.seasonResolved, true);
});

test('falls back to air order when the sequel chain is broken', async () => {
  const { execute } = createExecutor(OREGAIRU_SEARCH, {});
  const result = await resolveAnilistSeasonMedia(
    { title: 'My Teen Romantic Comedy SNAFU', season: 3, episode: 1 },
    { execute },
  );

  assert.equal(result?.id, 108489);
  assert.equal(result?.via, 'air-order');
  assert.equal(result?.seasonResolved, true);
});

test('reports seasonResolved false rather than silently using season 1', async () => {
  const { execute } = createExecutor([OREGAIRU_SEARCH[0]!], {});
  const result = await resolveAnilistSeasonMedia(
    { title: 'My Teen Romantic Comedy SNAFU', season: 3, episode: 1 },
    { execute },
  );

  assert.equal(result?.id, 14813);
  assert.equal(result?.seasonResolved, false);
  assert.equal(result?.via, 'anchor');
  assert.equal(result?.requestedSeason, 3);
});

test('does not filter the anchor by an episode count the season 1 entry cannot have', async () => {
  const { execute } = createExecutor(OREGAIRU_SEARCH, OREGAIRU_RELATIONS);
  // Season 2 episode 13 exceeds nothing, but season 3 episode 13 exceeds Climax!'s 12.
  const result = await resolveAnilistSeasonMedia(
    { title: 'My Teen Romantic Comedy SNAFU', season: 2, episode: 13 },
    { execute },
  );

  assert.equal(result?.id, 20698);
  assert.equal(result?.seasonResolved, true);
});

test('prefers the TV sequel when a franchise branches into other formats', async () => {
  const branching: Record<number, Array<{ relationType: string; node: AnilistSeasonMedia }>> = {
    1: [
      {
        relationType: 'SEQUEL',
        node: {
          id: 3,
          episodes: 6,
          format: 'ONA',
          seasonYear: 2019,
          title: { english: 'Spinoff' },
        },
      },
      {
        relationType: 'SEQUEL',
        node: { id: 2, episodes: 12, format: 'TV', seasonYear: 2020, title: { english: 'Show 2' } },
      },
    ],
  };
  const { execute } = createExecutor(
    [{ id: 1, episodes: 12, format: 'TV', seasonYear: 2018, title: { english: 'Show' } }],
    branching,
  );
  const result = await resolveAnilistSeasonMedia({ title: 'Show', season: 2 }, { execute });

  assert.equal(result?.id, 2);
  assert.equal(result?.via, 'sequel-chain');
});

test('air-order fallback refuses when the anchor is not the earliest entry', () => {
  const anchor: AnilistSeasonMedia = {
    id: 2,
    format: 'TV',
    seasonYear: 2020,
    title: { english: 'Show Later' },
  };
  const picked = pickByAirOrder(anchor, 2, [
    anchor,
    { id: 1, format: 'TV', seasonYear: 2015, title: { english: 'Show Earlier' } },
  ]);

  assert.equal(picked, null);
});

test('returns null when the search yields nothing', async () => {
  const { execute } = createExecutor([]);
  const result = await resolveAnilistSeasonMedia({ title: 'Nothing', season: 2 }, { execute });

  assert.equal(result, null);
});

test('refuses a season beyond the sequel-hop cap instead of returning a partial walk', async () => {
  // A 30-entry sequel chain: hopping the capped number of times would land on the wrong
  // season and report it resolved, so the walk must decline outright.
  const chain: Record<number, Array<{ relationType: string; node: AnilistSeasonMedia }>> = {};
  for (let id = 1; id < 30; id += 1) {
    chain[id] = [
      {
        relationType: 'SEQUEL',
        node: {
          id: id + 1,
          episodes: 12,
          format: 'TV',
          seasonYear: 2000 + id,
          title: { english: `Show ${id + 1}` },
        },
      },
    ];
  }
  const { execute, relationLookups } = createExecutor(
    [{ id: 1, episodes: 12, format: 'TV', seasonYear: 2000, title: { english: 'Show' } }],
    chain,
  );

  const result = await resolveAnilistSeasonMedia({ title: 'Show', season: 25 }, { execute });

  assert.equal(result?.id, 1);
  assert.equal(result?.seasonResolved, false);
  assert.equal(result?.via, 'anchor');
  // Declines before spending any relation requests.
  assert.deepEqual(relationLookups, []);
});

test('air-order fallback declines when a franchise entry has no air year', async () => {
  // The middle season has no year, so ordering it would shift every later season by one.
  const { execute } = createExecutor(
    [
      { id: 1, episodes: 12, format: 'TV', seasonYear: 2013, title: { english: 'Show' } },
      { id: 2, episodes: 12, format: 'TV', title: { english: 'Show Zoku' } },
      { id: 3, episodes: 12, format: 'TV', seasonYear: 2020, title: { english: 'Show Kan' } },
    ],
    {},
  );

  const result = await resolveAnilistSeasonMedia({ title: 'Show', season: 2 }, { execute });

  assert.equal(result?.id, 1);
  assert.equal(result?.seasonResolved, false);
  assert.equal(result?.via, 'anchor');
});
