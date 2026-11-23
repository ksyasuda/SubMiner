import test from 'node:test';
import assert from 'node:assert/strict';

import {
  ANILIST_SEASON_RELATIONS_QUERY,
  ANILIST_SEASON_SEARCH_QUERY,
  namedSeasons,
  pickByAirOrder,
  resolveAnilistSeasonMedia,
  splitPartMarker,
  stripSeasonSuffix,
  type AnilistSeasonMedia,
} from './season-resolver';
import {
  AOT_CHAIN,
  BLEACH_TYBW_CHAIN,
  DR_STONE_CHAIN,
  JOJO_CHAIN,
  MUSHOKU_CHAIN,
  OVERLORD_CHAIN,
  REZERO_CHAIN,
} from './__fixtures__/anilist-season-chains';

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

test('AniList refuses URL-derived search titles before making a request', async () => {
  let requests = 0;
  for (const title of [
    'https://example.com/stream?api_key=test-secret',
    'stream?static=true&api_key=test-secret',
    'stream static true api key test secret',
  ]) {
    const result = await resolveAnilistSeasonMedia(
      { title },
      {
        execute: async () => {
          requests += 1;
          throw new Error('must not send URL-derived searches');
        },
      },
    );
    assert.equal(result, null);
  }
  assert.equal(requests, 0);
});

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
  // The last lookup checks one season past the count for an entry naming season 3.
  assert.deepEqual(relationLookups, [14813, 20698, 108489]);
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
  assert.equal(result?.via, 'named-season');
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
  assert.equal(result?.via, 'named-season');
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

/** Executor over a fixture chain: search returns the whole chain, entry i's SEQUEL is i + 1. */
function createChainExecutor(chain: AnilistSeasonMedia[]) {
  const relations: Record<number, Array<{ relationType: string; node: AnilistSeasonMedia }>> = {};
  chain.forEach((entry, index) => {
    const next = chain[index + 1];
    relations[entry.id] = next ? [{ relationType: 'SEQUEL', node: next }] : [];
  });
  return createExecutor(chain, relations);
}

async function resolveInChain(
  chain: AnilistSeasonMedia[],
  input: { title: string; season?: number | null; episode?: number | null },
) {
  const executor = createChainExecutor(chain);
  const result = await resolveAnilistSeasonMedia(input, { execute: executor.execute });
  return { result, relationLookups: executor.relationLookups };
}

test('split-cour back halves do not count as seasons', async () => {
  const title = 'Mushoku Tensei: Jobless Reincarnation';
  const s2 = await resolveInChain(MUSHOKU_CHAIN, { title, season: 2, episode: 1 });
  assert.equal(s2.result?.id, 146065);
  assert.equal(s2.result?.seasonResolved, true);

  const s3 = await resolveInChain(MUSHOKU_CHAIN, { title, season: 3, episode: 1 });
  assert.equal(s3.result?.id, 178789);
  // Four hops for two seasons: "Cour 2" and "Season 2 Part 2" were walked through.
  assert.deepEqual(s3.relationLookups, [108465, 127720, 146065, 166873]);

  const reZeroS3 = await resolveInChain(REZERO_CHAIN, { title: 'Re:ZERO', season: 3 });
  assert.equal(reZeroS3.result?.id, 163134);
  const reZeroS4 = await resolveInChain(REZERO_CHAIN, { title: 'Re:ZERO', season: 4 });
  assert.equal(reZeroS4.result?.id, 189046);
});

test('a special in the sequel chain is walked through without counting', async () => {
  const s3 = await resolveInChain(DR_STONE_CHAIN, { title: 'Dr. STONE', season: 3 });
  assert.equal(s3.result?.id, 131518);
  const s4 = await resolveInChain(DR_STONE_CHAIN, { title: 'Dr. STONE', season: 4 });
  assert.equal(s4.result?.id, 172019);
});

test('season names ignore part and cour numbers, including localized ones', () => {
  // "Cour 2" carries the Thai synonym "... พาร์ท 2" ("part 2").
  assert.deepEqual([...namedSeasons(MUSHOKU_CHAIN[1]!)], []);
  assert.deepEqual([...namedSeasons(MUSHOKU_CHAIN[3]!)], [2]);
  assert.deepEqual([...namedSeasons(OVERLORD_CHAIN[1]!)], [2]);
  assert.deepEqual([...namedSeasons(DR_STONE_CHAIN[1]!)], [2]);
  assert.deepEqual([...namedSeasons(AOT_CHAIN[4]!)], [4]);
});

test('an entry naming the requested season beats the hop count', async () => {
  // The second cour has no marker anywhere, so counting alone would call it season 2.
  const chain: AnilistSeasonMedia[] = [
    { id: 1, episodes: 12, format: 'TV', seasonYear: 2020, title: { english: 'Show' } },
    { id: 2, episodes: 12, format: 'TV', seasonYear: 2020, title: { english: 'Show: Arc Two' } },
    { id: 3, episodes: 12, format: 'TV', seasonYear: 2022, title: { english: 'Show Season 2' } },
  ];
  const { result } = await resolveInChain(chain, { title: 'Show', season: 2 });
  assert.equal(result?.id, 3);
  assert.equal(result?.via, 'named-season');

  const overlord = await resolveInChain(OVERLORD_CHAIN, { title: 'Overlord', season: 2 });
  assert.equal(overlord.result?.id, 98437);
  assert.equal(overlord.result?.via, 'named-season');
});

test('a counted entry whose title names another season is not accepted', async () => {
  const chain: AnilistSeasonMedia[] = [
    { id: 1, episodes: 12, format: 'TV', seasonYear: 2020, title: { english: 'Show' } },
    { id: 2, episodes: 12, format: 'TV', seasonYear: 2021, title: { english: 'Show: Arc Two' } },
    { id: 3, episodes: 12, format: 'TV', seasonYear: 2022, title: { english: 'Show Season 2' } },
  ];
  // Counting lands on id 3 for season 3, but it says it is season 2.
  const { result } = await resolveInChain(chain, { title: 'Show', season: 3 });
  assert.equal(result?.seasonResolved, false);
});

test('JoJo parts are seasons, not split-cour continuations', async () => {
  assert.equal(splitPartMarker('JoJo no Kimyou na Bouken Part 5: Ougon no Kaze'), null);
  assert.deepEqual(splitPartMarker("JoJo's Bizarre Adventure Part 5"), {
    base: 'jojo s bizarre adventure',
    part: 5,
  });
  // Golden Wind ends in "Part 5" and Stone Ocean in "Part 6": same base, different
  // parts, so Stone Ocean must still count as the next season.
  const { result } = await resolveInChain(JOJO_CHAIN, {
    title: "JoJo's Bizarre Adventure: Diamond is Unbreakable",
    season: 3,
  });
  assert.equal(result?.id, 131942);
  assert.equal(result?.via, 'sequel-chain');
});

test('Attack on Titan resolves its final season past the season 3 back half', async () => {
  const { result } = await resolveInChain(AOT_CHAIN, { title: 'Attack on Titan', season: 4 });
  assert.equal(result?.id, 110277);
});

test('air-order fallback drops back halves but keeps a season spelled "Part 1"', () => {
  // AoT season 3 carries the synonym "ผ่าพิภพไททัน ภาค 3 Part 1".
  assert.equal(pickByAirOrder(AOT_CHAIN[0]!, 3, AOT_CHAIN)?.id, 99147);
  assert.equal(pickByAirOrder(AOT_CHAIN[0]!, 4, AOT_CHAIN)?.id, 110277);
  assert.equal(pickByAirOrder(MUSHOKU_CHAIN[0]!, 3, MUSHOKU_CHAIN)?.id, 178789);
  // Only three real seasons have aired, so a fourth must not fall out of the list.
  assert.equal(pickByAirOrder(MUSHOKU_CHAIN[0]!, 4, MUSHOKU_CHAIN.slice(0, 5)), null);
});

test('a TYBW cour marked only in a differently spelled synonym is not season 2', async () => {
  const { result } = await resolveInChain(BLEACH_TYBW_CHAIN, {
    title: 'BLEACH: Thousand-Year Blood War',
    season: 2,
  });
  // Every sequel is a later cour of the same season, so there is no season 2 to find.
  assert.equal(result?.seasonResolved, false);
  assert.equal(result?.id, 116674);
});

test('an episode past the season entry lands on the split-cour part that airs it', async () => {
  const mushoku = await resolveInChain(MUSHOKU_CHAIN, {
    title: 'Mushoku Tensei: Jobless Reincarnation',
    season: 1,
    episode: 15,
  });
  assert.equal(mushoku.result?.id, 108465);
  assert.equal(mushoku.result?.episodeEntry?.id, 127720);
  assert.equal(mushoku.result?.episodeEntry?.episode, 4);

  const mushokuS2 = await resolveInChain(MUSHOKU_CHAIN, {
    title: 'Mushoku Tensei: Jobless Reincarnation',
    season: 2,
    episode: 20,
  });
  assert.equal(mushokuS2.result?.id, 146065);
  assert.equal(mushokuS2.result?.episodeEntry?.id, 166873);
  assert.equal(mushokuS2.result?.episodeEntry?.episode, 7);

  // Three cours of 13 + 13 + 14: episode 28 is the third cour's second.
  const bleach = await resolveInChain(BLEACH_TYBW_CHAIN, {
    title: 'BLEACH: Thousand-Year Blood War',
    episode: 28,
  });
  assert.equal(bleach.result?.id, 116674);
  assert.equal(bleach.result?.episodeEntry?.id, 169755);
  assert.equal(bleach.result?.episodeEntry?.episode, 2);
});

test('an episode inside the season entry needs no remap or relation lookup', async () => {
  const { result, relationLookups } = await resolveInChain(MUSHOKU_CHAIN, {
    title: 'Mushoku Tensei: Jobless Reincarnation',
    season: 1,
    episode: 5,
  });
  assert.equal(result?.id, 108465);
  assert.equal(result?.episodeEntry, null);
  assert.deepEqual(relationLookups, []);
});

test('an episode past a season with no continuation is not moved into the next season', async () => {
  const { result } = await resolveInChain(OVERLORD_CHAIN, {
    title: 'Overlord',
    season: 1,
    episode: 20,
  });
  assert.equal(result?.id, 20832);
  assert.equal(result?.episodeEntry, null);
});

test('a sequel naming its own season is not a back half, even with a part synonym', async () => {
  const sequel: AnilistSeasonMedia = {
    id: 2,
    episodes: 12,
    format: 'TV',
    seasonYear: 2022,
    title: { english: 'Part Synonym Show Season 2' },
    synonyms: ['Part Synonym Show Part 2'],
  };
  // The anchor names season 1 itself...
  const named = await resolveInChain(
    [
      {
        id: 1,
        episodes: 12,
        format: 'TV',
        seasonYear: 2020,
        title: { english: 'Part Synonym Show' },
        synonyms: ['Part Synonym Show Season 1'],
      },
      sequel,
    ],
    { title: 'Part Synonym Show', season: 2 },
  );
  assert.equal(named.result?.id, 2);
  // ...or is season 1 because the walk starts there.
  const unnamed = await resolveInChain(
    [
      {
        id: 1,
        episodes: 12,
        format: 'TV',
        seasonYear: 2020,
        title: { english: 'Part Synonym Show' },
      },
      sequel,
    ],
    { title: 'Part Synonym Show', season: 2 },
  );
  assert.equal(unnamed.result?.id, 2);
});

test('air order keeps a sequel naming its own season despite a part synonym', () => {
  const anchor: AnilistSeasonMedia = {
    id: 1,
    episodes: 12,
    format: 'TV',
    seasonYear: 2020,
    title: { english: 'Air Order Show' },
  };
  const sequel: AnilistSeasonMedia = {
    id: 2,
    episodes: 12,
    format: 'TV',
    seasonYear: 2022,
    title: { english: 'Air Order Show Season 2' },
    synonyms: ['Air Order Show Part 2'],
  };
  assert.equal(pickByAirOrder(anchor, 2, [anchor, sequel])?.id, 2);
});

test('a season 1 episode past its length does not move to a longer later season', async () => {
  const { result } = await resolveInChain(
    [
      {
        id: 1,
        episodes: 12,
        format: 'TV',
        seasonYear: 2020,
        title: { english: 'Long Later Show' },
      },
      {
        id: 2,
        episodes: 24,
        format: 'TV',
        seasonYear: 2022,
        title: { english: 'Long Later Show Season 2' },
      },
    ],
    { title: 'Long Later Show', season: 1, episode: 20 },
  );
  assert.equal(result?.id, 1);
  assert.equal(result?.episodeEntry, null);
});

test('a sequel that repeats its predecessor name is still a new season', async () => {
  // No part marker, so the identical title must not be read as a continuation.
  const chain: AnilistSeasonMedia[] = [
    { id: 1, episodes: 12, format: 'TV', seasonYear: 2020, title: { english: 'Same Name Show' } },
    { id: 2, episodes: 12, format: 'TV', seasonYear: 2022, title: { english: 'Same Name Show' } },
  ];
  const { result } = await resolveInChain(chain, { title: 'Same Name Show', season: 2 });
  assert.equal(result?.id, 2);
  assert.equal(result?.seasonResolved, true);
});

test('a rate limit mid-walk is surfaced instead of answered from air order', async () => {
  const { execute } = createChainExecutor(MUSHOKU_CHAIN);
  let relationCalls = 0;
  const failing = async <T>(query: string, variables: Record<string, unknown>): Promise<T> => {
    if (query === ANILIST_SEASON_RELATIONS_QUERY && ++relationCalls === 2) {
      throw new Error('AniList rate limited (429)');
    }
    return execute<T>(query, variables);
  };
  // Air order alone would answer season 3 correctly from this search result, but it must
  // not answer from a franchise picture the failed walk never confirmed.
  await assert.rejects(
    resolveAnilistSeasonMedia(
      { title: 'Mushoku Tensei: Jobless Reincarnation', season: 3 },
      { execute: failing },
    ),
    /429/,
  );
});
