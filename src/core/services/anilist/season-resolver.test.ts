import test from 'node:test';
import assert from 'node:assert/strict';

import {
  ANILIST_SEASON_RELATIONS_QUERY,
  ANILIST_SEASON_SEARCH_QUERY,
  pickByAirOrder,
  resolveAnilistSeasonMedia,
  splitPartMarker,
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

test('a sequel resolution is not certified by the anchor exact-title evidence', async () => {
  // The anchor matched the search title exactly, but the hopped-to entry is a
  // different inference (a split-cour chain can land one season short), so the
  // sequel result must report its own title evidence, not the anchor's.
  const { execute } = createExecutor(OREGAIRU_SEARCH, OREGAIRU_RELATIONS);
  const result = await resolveAnilistSeasonMedia(
    { title: 'My Teen Romantic Comedy SNAFU', season: 2, episode: 1 },
    { execute },
  );

  assert.equal(result?.id, 20698);
  assert.equal(result?.via, 'sequel-chain');
  assert.equal(result?.exactTitleMatch, false);
});

test('a sequel resolution whose own title matches the parsed title stays exact', async () => {
  const anchor: AnilistSeasonMedia = {
    id: 1,
    episodes: 12,
    format: 'TV',
    title: { english: 'Show' },
  };
  const sequel: AnilistSeasonMedia = {
    id: 2,
    episodes: 12,
    format: 'TV',
    title: { english: 'Show 2nd Season' },
  };
  const { execute } = createExecutor([anchor], {
    1: [{ relationType: 'SEQUEL', node: sequel }],
  });

  const result = await resolveAnilistSeasonMedia(
    { title: 'Show 2nd Season', season: 2, episode: 1 },
    { execute },
  );

  assert.equal(result?.id, 2);
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

/**
 * Real AniList payloads for Mushoku Tensei, whose sequel chain interleaves
 * split-cour continuations with real seasons:
 * S1 -> "Cour 2" -> S2 -> "Season 2 Part 2" -> S3.
 */
const MUSHOKU_SEARCH: AnilistSeasonMedia[] = [
  {
    id: 108465,
    episodes: 11,
    format: 'TV',
    seasonYear: 2021,
    title: {
      romaji: 'Mushoku Tensei: Isekai Ittara Honki Dasu',
      english: 'Mushoku Tensei: Jobless Reincarnation',
      native: null,
    },
  },
  {
    id: 127720,
    episodes: 12,
    format: 'TV',
    seasonYear: 2021,
    title: {
      romaji: 'Mushoku Tensei: Isekai Ittara Honki Dasu Part 2',
      english: 'Mushoku Tensei: Jobless Reincarnation Cour 2',
      native: null,
    },
  },
  {
    id: 146065,
    episodes: 13,
    format: 'TV',
    seasonYear: 2023,
    title: {
      romaji: 'Mushoku Tensei II: Isekai Ittara Honki Dasu',
      english: 'Mushoku Tensei: Jobless Reincarnation Season 2',
      native: null,
    },
  },
  {
    id: 166873,
    episodes: 12,
    format: 'TV',
    seasonYear: 2024,
    title: {
      romaji: 'Mushoku Tensei II: Isekai Ittara Honki Dasu Part 2',
      english: 'Mushoku Tensei: Jobless Reincarnation Season 2 Part 2',
      native: null,
    },
  },
  {
    id: 178789,
    episodes: 14,
    format: 'TV',
    seasonYear: 2026,
    title: {
      romaji: 'Mushoku Tensei III: Isekai Ittara Honki Dasu',
      english: 'Mushoku Tensei: Jobless Reincarnation Season 3',
      native: null,
    },
  },
];

const MUSHOKU_RELATIONS: Record<
  number,
  Array<{ relationType: string; node: AnilistSeasonMedia }>
> = {
  108465: [{ relationType: 'SEQUEL', node: MUSHOKU_SEARCH[1]! }],
  127720: [{ relationType: 'SEQUEL', node: MUSHOKU_SEARCH[2]! }],
  146065: [{ relationType: 'SEQUEL', node: MUSHOKU_SEARCH[3]! }],
  166873: [{ relationType: 'SEQUEL', node: MUSHOKU_SEARCH[4]! }],
};

test('splitPartMarker separates a split-cour marker from the season title', () => {
  assert.deepEqual(splitPartMarker('Mushoku Tensei: Jobless Reincarnation Season 2 Part 2'), {
    base: 'mushoku tensei: jobless reincarnation season 2',
    hasPart: true,
  });
  assert.deepEqual(splitPartMarker('Mushoku Tensei: Jobless Reincarnation Cour 2'), {
    base: 'mushoku tensei: jobless reincarnation',
    hasPart: true,
  });
  assert.deepEqual(splitPartMarker('Attack on Titan Final Season Part 3'), {
    base: 'attack on titan final season',
    hasPart: true,
  });
  assert.deepEqual(splitPartMarker('進撃の巨人 第2部'), { base: '進撃の巨人', hasPart: true });
});

test('splitPartMarker leaves a numbered work alone', () => {
  // The number is followed by a subtitle, so it names the work, not a cour.
  assert.deepEqual(splitPartMarker('JoJo no Kimyou na Bouken Part 5: Ougon no Kaze'), {
    base: 'jojo no kimyou na bouken part 5: ougon no kaze',
    hasPart: false,
  });
  assert.deepEqual(splitPartMarker('Mushoku Tensei: Jobless Reincarnation Season 3'), {
    base: 'mushoku tensei: jobless reincarnation season 3',
    hasPart: false,
  });
});

test('a split-cour continuation does not consume a season step', async () => {
  const { execute, relationLookups } = createExecutor(MUSHOKU_SEARCH, MUSHOKU_RELATIONS);
  const result = await resolveAnilistSeasonMedia(
    { title: 'Mushoku Tensei: Jobless Reincarnation', season: 3, episode: 4 },
    { execute },
  );

  assert.equal(result?.id, 178789);
  assert.equal(result?.title, 'Mushoku Tensei: Jobless Reincarnation Season 3');
  assert.equal(result?.seasonResolved, true);
  assert.equal(result?.via, 'sequel-chain');
  // Four hops for three seasons: two of them were cour continuations.
  assert.deepEqual(relationLookups, [108465, 127720, 146065, 166873]);
});

test('season 2 stops at the front half rather than its second cour', async () => {
  const { execute } = createExecutor(MUSHOKU_SEARCH, MUSHOKU_RELATIONS);
  const result = await resolveAnilistSeasonMedia(
    { title: 'Mushoku Tensei: Jobless Reincarnation', season: 2, episode: 1 },
    { execute },
  );

  assert.equal(result?.id, 146065);
  assert.equal(result?.seasonResolved, true);
});

test('air-order fallback skips split-cour entries too', () => {
  const anchor = MUSHOKU_SEARCH[0]!;
  assert.equal(pickByAirOrder(anchor, 2, MUSHOKU_SEARCH)?.id, 146065);
  assert.equal(pickByAirOrder(anchor, 3, MUSHOKU_SEARCH)?.id, 178789);
  // Only three real seasons exist, so a fourth must not fall out of the list.
  assert.equal(pickByAirOrder(anchor, 4, MUSHOKU_SEARCH), null);
});

test('a sequel that repeats its predecessor name is still a new season', async () => {
  // No part marker, so the identical title must not be read as a continuation.
  const search: AnilistSeasonMedia[] = [
    {
      id: 1,
      episodes: 12,
      format: 'TV',
      seasonYear: 2020,
      title: { romaji: 'Same Name Show', english: 'Same Name Show', native: null },
    },
    {
      id: 2,
      episodes: 12,
      format: 'TV',
      seasonYear: 2022,
      title: { romaji: 'Same Name Show', english: 'Same Name Show', native: null },
    },
  ];
  const { execute } = createExecutor(search, {
    1: [{ relationType: 'SEQUEL', node: search[1]! }],
  });
  const result = await resolveAnilistSeasonMedia(
    { title: 'Same Name Show', season: 2, episode: 3 },
    { execute },
  );

  assert.equal(result?.id, 2);
  assert.equal(result?.seasonResolved, true);
});

/**
 * Real AniList payloads for Bleach TYBW, whose cours are titled with their own
 * arc names — the part markers live only in localized synonyms, and the English
 * title hyphenates "Thousand-Year" while those synonyms do not.
 */
const BLEACH_SEARCH: AnilistSeasonMedia[] = [
  {
    id: 116674,
    episodes: 13,
    format: 'TV',
    seasonYear: 2022,
    synonyms: ['BLEACH TYBW'],
    title: {
      romaji: 'BLEACH: Sennen Kessen-hen',
      english: 'BLEACH: Thousand-Year Blood War',
      native: 'BLEACH 千年血戦篇',
    },
  },
  {
    id: 159322,
    episodes: 13,
    format: 'TV',
    seasonYear: 2023,
    synonyms: [
      'BLEACH: Thousand Year Blood War Part 2',
      'BLEACH 千年血戦篇 第2クール',
      'BLEACH TYBW',
    ],
    title: {
      romaji: 'BLEACH: Sennen Kessen-hen - Ketsubetsu-tan',
      english: 'BLEACH: Thousand-Year Blood War - The Separation',
      native: 'BLEACH 千年血戦篇-訣別譚-',
    },
  },
];

test('a cour marked only in a synonym, and spelled without the hyphen, still counts as one', async () => {
  const { execute } = createExecutor(BLEACH_SEARCH, {
    116674: [{ relationType: 'SEQUEL', node: BLEACH_SEARCH[1]! }],
  });
  const result = await resolveAnilistSeasonMedia(
    { title: 'BLEACH: Thousand-Year Blood War', season: 2, episode: 1 },
    { execute },
  );

  // The only sequel is cour 2 of the same season, so there is no season 2 to
  // find — reporting that beats confidently returning the wrong arc.
  assert.equal(result?.seasonResolved, false);
  assert.notEqual(result?.id, 159322);
});

test('splitPartMarker reads the Japanese cour marker', () => {
  assert.deepEqual(splitPartMarker('BLEACH 千年血戦篇 第2クール'), {
    base: 'bleach 千年血戦篇',
    hasPart: true,
  });
});

/** Dr. STONE routes its sequel chain through a one-episode special. */
const DR_STONE_SEARCH: AnilistSeasonMedia[] = [
  {
    id: 105333,
    episodes: 24,
    format: 'TV',
    seasonYear: 2019,
    title: { romaji: 'Dr. STONE', english: 'Dr. STONE', native: null },
  },
  {
    id: 113936,
    episodes: 11,
    format: 'TV',
    seasonYear: 2021,
    title: { romaji: 'Dr. STONE: STONE WARS', english: 'Dr. STONE: STONE WARS', native: null },
  },
  {
    id: 142876,
    episodes: 1,
    format: 'SPECIAL',
    seasonYear: 2022,
    title: {
      romaji: 'Dr. STONE: Ryuusui',
      english: 'Dr. STONE Special Episode – RYUSUI',
      native: null,
    },
  },
  {
    id: 131518,
    episodes: 11,
    format: 'TV',
    seasonYear: 2023,
    title: { romaji: 'Dr. STONE: NEW WORLD', english: 'Dr. STONE New World', native: null },
  },
];

test('a special in the sequel chain is walked through but does not count as a season', async () => {
  const { execute } = createExecutor(DR_STONE_SEARCH, {
    105333: [{ relationType: 'SEQUEL', node: DR_STONE_SEARCH[1]! }],
    113936: [{ relationType: 'SEQUEL', node: DR_STONE_SEARCH[2]! }],
    142876: [{ relationType: 'SEQUEL', node: DR_STONE_SEARCH[3]! }],
  });
  const result = await resolveAnilistSeasonMedia(
    { title: 'Dr. STONE', season: 3, episode: 1 },
    { execute },
  );

  assert.equal(result?.id, 131518);
  assert.equal(result?.title, 'Dr. STONE New World');
  assert.equal(result?.seasonResolved, true);
});

test('air-order fallback keeps a season whose own synonym spells it as a part', () => {
  // AniList carries "…Season 3 Part 1" as a localized synonym *of season 3*.
  const search: AnilistSeasonMedia[] = [
    {
      id: 1,
      episodes: 25,
      format: 'TV',
      seasonYear: 2013,
      title: { romaji: 'Some Titans', english: 'Some Titans', native: null },
    },
    {
      id: 2,
      episodes: 12,
      format: 'TV',
      seasonYear: 2017,
      title: { romaji: 'Some Titans Season 2', english: 'Some Titans Season 2', native: null },
    },
    {
      id: 3,
      episodes: 12,
      format: 'TV',
      seasonYear: 2018,
      synonyms: ['Some Titans Season 3 Part 1'],
      title: { romaji: 'Some Titans Season 3', english: 'Some Titans Season 3', native: null },
    },
  ];

  assert.equal(pickByAirOrder(search[0]!, 3, search)?.id, 3);
});

test('a transport failure mid-walk is surfaced instead of guessed around', async () => {
  const { execute } = createExecutor(MUSHOKU_SEARCH, MUSHOKU_RELATIONS);
  const failing = async <T>(query: string, variables: Record<string, unknown>): Promise<T> => {
    if (query === ANILIST_SEASON_RELATIONS_QUERY) {
      throw new Error('Too Many Requests.');
    }
    return execute<T>(query, variables);
  };

  // Air order could answer here, but it would be answering from a franchise
  // picture the failed walk never confirmed.
  await assert.rejects(
    () =>
      resolveAnilistSeasonMedia(
        { title: 'Mushoku Tensei: Jobless Reincarnation', season: 3, episode: 4 },
        { execute: failing },
      ),
    /Too Many Requests/,
  );
});
