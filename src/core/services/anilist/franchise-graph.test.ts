import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildFranchiseGraph,
  walkFranchiseRelations,
  type AnilistMediaNode,
  type FranchiseWalkOptions,
} from './franchise-graph';

function entry(
  anilistId: number,
  format: string | null,
  startDate: number | null,
  relations: Array<[string, number]> = [],
): AnilistMediaNode {
  return {
    anilistId,
    format,
    titleRomaji: `romaji ${anilistId}`,
    titleEnglish: null,
    titleNative: null,
    startDate,
    episodes: null,
    relations: relations.map(([relationType, targetId]) => ({ relationType, targetId })),
    fetchedAtMs: 0,
  };
}

// Overlord: two recap films sit between seasons one and two; Isekai Quartet only shares characters.
const overlord = [
  entry(20832, 'TV', 20150707, [
    ['SEQUEL', 98437],
    ['SUMMARY', 98873],
    ['CHARACTER', 104454],
  ]),
  entry(98873, 'MOVIE', 20170225, [
    ['PARENT', 20832],
    ['SEQUEL', 98874],
  ]),
  entry(98874, 'MOVIE', 20170311, [
    ['PREQUEL', 98873],
    ['SEQUEL', 98437],
  ]),
  entry(98437, 'TV', 20180109, [
    ['PREQUEL', 20832],
    ['PREQUEL', 98874],
  ]),
  entry(104454, 'TV_SHORT', 20190410, [['CHARACTER', 20832]]),
];

test('seasons, films and recaps share a root named after the first TV series', () => {
  const graph = buildFranchiseGraph(overlord);
  for (const id of [20832, 98873, 98874, 98437]) assert.equal(graph.rootOf(id), 20832);
  assert.equal(graph.rootOf(104454), 104454);
  assert.equal(graph.rootOf(1), undefined);
});

test('excluded relation types keep franchises apart even when the other end says PARENT', () => {
  const graph = buildFranchiseGraph([
    entry(1, 'TV', 20060101, [
      ['SPIN_OFF', 2],
      ['ALTERNATIVE', 3],
      ['ADAPTATION', 4],
      ['OTHER', 5],
    ]),
    // A spin-off pointing back at its original as PARENT.
    entry(2, 'TV', 20130101, [['PARENT', 1]]),
    entry(3, 'TV', 20140101, [['ALTERNATIVE', 1]]),
  ]);
  assert.deepEqual(
    [1, 2, 3, 4, 5].map((id) => graph.rootOf(id)),
    [1, 2, 3, undefined, undefined],
  );
});

test('root falls back to episodic formats, then earliest entry, then lowest id', () => {
  const ona = buildFranchiseGraph([
    entry(10, 'MOVIE', 20000101, [['SEQUEL', 11]]),
    entry(11, 'ONA', 20050101, [['PREQUEL', 10]]),
  ]);
  assert.equal(ona.rootOf(10), 11);
  const films = buildFranchiseGraph([
    entry(20, 'MOVIE', 20100101, [['PREQUEL', 21]]),
    entry(21, 'MOVIE', 20090101, [['SEQUEL', 20]]),
    entry(22, 'OVA', null, [['PARENT', 20]]),
  ]);
  assert.equal(films.rootOf(22), 21);
});

test('owned films join through an unowned, uncached TV entry they both point at', () => {
  const graph = buildFranchiseGraph([
    entry(528, 'MOVIE', 19980718, [['PARENT', 527]]),
    entry(1526, 'MOVIE', 20050716, [['PARENT', 527]]),
  ]);
  assert.equal(graph.rootOf(528), graph.rootOf(1526));
});

test('relation cycles terminate and form one franchise', () => {
  const graph = buildFranchiseGraph([
    entry(1, 'TV', 20000101, [['SEQUEL', 2]]),
    entry(2, 'TV', 20010101, [['SEQUEL', 3]]),
    entry(3, 'TV', 20020101, [['SEQUEL', 1]]),
  ]);
  assert.deepEqual(
    [1, 2, 3].map((id) => graph.rootOf(id)),
    [1, 1, 1],
  );
});

function walkHarness(all: AnilistMediaNode[], overrides: Partial<FranchiseWalkOptions> = {}) {
  const asked: number[][] = [];
  const known = new Map<number, AnilistMediaNode>();
  const options: FranchiseWalkOptions = {
    seeds: [],
    known,
    isFresh: () => true,
    fetchNodes: async (ids) => {
      asked.push([...ids].sort((a, b) => a - b));
      return all.filter((node) => ids.includes(node.anilistId));
    },
    onBatch: () => {},
    now: () => 5,
    maxDepth: 6,
    maxFetch: 500,
    ...overrides,
  };
  return { asked, known, options };
}

test('walk follows franchise links through entries the user does not own', async () => {
  const pokemon = [
    entry(528, 'MOVIE', 19980718, [['PARENT', 527]]),
    entry(527, 'TV', 19970401, [
      ['SIDE_STORY', 528],
      ['SEQUEL', 1564],
    ]),
    entry(1564, 'TV', 20021121, [
      ['PREQUEL', 527],
      ['SIDE_STORY', 1526],
      ['OTHER', 9],
    ]),
    entry(1526, 'MOVIE', 20050716, [['PARENT', 1564]]),
    entry(9, 'TV', 20000101),
  ];
  const { asked, known, options } = walkHarness(pokemon, { seeds: [528, 1526, 404] });
  await walkFranchiseRelations(options);
  assert.deepEqual(asked, [
    [404, 528, 1526],
    [527, 1564],
  ]);
  const graph = buildFranchiseGraph(known.values());
  assert.equal(graph.rootOf(528), 527);
  assert.equal(graph.rootOf(1526), 527);
  // Missing ids are remembered so they are not asked about again; OTHER links are not followed.
  assert.equal(known.get(404)?.format, null);
  assert.equal(known.has(9), false);
});

test('walk does not follow the links of an entry that vetoes how it was reached', async () => {
  const { asked, options } = walkHarness(
    [
      entry(2, 'TV', 20130101, [['PARENT', 1]]),
      entry(1, 'TV', 20060101, [
        ['SPIN_OFF', 2],
        ['SEQUEL', 3],
      ]),
      entry(3, 'TV', 20070101),
    ],
    { seeds: [2] },
  );
  await walkFranchiseRelations(options);
  assert.deepEqual(asked, [[2], [1]]);
});

test('walk respects the depth and fetch bounds', async () => {
  const chain = Array.from({ length: 10 }, (_, i) =>
    entry(i + 1, 'TV', 20000101 + i, i < 9 ? [['SEQUEL', i + 2]] : []),
  );
  const deep = walkHarness(chain, { seeds: [1], maxDepth: 3 });
  await walkFranchiseRelations(deep.options);
  assert.deepEqual([...deep.known.keys()], [1, 2, 3, 4]);

  const budget = walkHarness(chain, { seeds: [1], maxFetch: 2 });
  const result = await walkFranchiseRelations(budget.options);
  assert.deepEqual(result, { fetched: 2, truncated: true });
  assert.deepEqual([...budget.known.keys()], [1, 2]);
});

test('walk follows fresh cached entries without asking and keeps old data on a failed fetch', async () => {
  const cached = entry(1, 'TV', 20000101, [['SEQUEL', 2]]);
  const stale = { ...entry(2, 'TV', 20010101, [['PREQUEL', 1]]), fetchedAtMs: -1 };
  const { known, options } = walkHarness([], {
    seeds: [1],
    isFresh: (node) => node.fetchedAtMs >= 0,
    fetchNodes: async () => {
      throw new Error('AniList unavailable');
    },
  });
  known.set(1, cached);
  known.set(2, stale);
  await assert.rejects(walkFranchiseRelations(options), /AniList unavailable/);
  assert.deepEqual(known.get(2), stale);
});

test('an entry AniList stops returning keeps its cached links', async () => {
  const previous = { ...entry(2, 'TV', 20010101, [['PREQUEL', 1]]), fetchedAtMs: -1 };
  const { known, options } = walkHarness([], {
    seeds: [2],
    isFresh: (node) => node.fetchedAtMs >= 0,
  });
  known.set(2, previous);
  await walkFranchiseRelations(options);
  assert.deepEqual(known.get(2), { ...previous, fetchedAtMs: 5 });
});
