/**
 * Franchise grouping over AniList's relation graph. AniList keeps every
 * season, movie and OVA as its own entry and says how entries relate; the
 * stats library lists a franchise once by following those links, including
 * through entries the user does not own (two movies meet at a TV series that
 * is not in the library). Nothing here touches the database or the network.
 */

/**
 * Relation types that make two entries parts of one franchise. SPIN_OFF,
 * ALTERNATIVE, CHARACTER, OTHER, ADAPTATION and SOURCE are left out on purpose:
 * following them collapses Gundam or Fate into a single card.
 */
export const FRANCHISE_RELATION_TYPES: ReadonlySet<string> = new Set([
  'PREQUEL',
  'SEQUEL',
  'PARENT',
  'SIDE_STORY',
  'SUMMARY',
  'COMPILATION',
  'CONTAINS',
]);

export interface AnilistRelationEdge {
  relationType: string;
  targetId: number;
}

/** One cached AniList entry and its links to other anime entries. */
export interface AnilistMediaNode {
  anilistId: number;
  /** TV, MOVIE, OVA...; null when AniList did not return the entry. */
  format: string | null;
  titleRomaji: string | null;
  titleEnglish: string | null;
  titleNative: string | null;
  /** First air date as yyyymmdd (month/day 0 when unknown); null when AniList has no year. */
  startDate: number | null;
  episodes: number | null;
  relations: AnilistRelationEdge[];
  fetchedAtMs: number;
}

export interface FranchiseGraph {
  node(anilistId: number): AnilistMediaNode | undefined;
  /** AniList id the franchise containing `anilistId` is named after; undefined when unknown. */
  rootOf(anilistId: number): number | undefined;
}

export function anilistNodeTitle(node: AnilistMediaNode): string | null {
  return node.titleEnglish ?? node.titleRomaji ?? node.titleNative;
}

/**
 * True when `node` names `targetId` only through relations that do not join a
 * franchise. AniList records most links from both ends; a spin-off can point
 * back at its original as PARENT, so the original's SPIN_OFF edge has to veto
 * the join or excluding SPIN_OFF would not keep the two apart.
 */
function vetoes(node: AnilistMediaNode, targetId: number): boolean {
  const edges = node.relations.filter((edge) => edge.targetId === targetId);
  return edges.length > 0 && !edges.some((edge) => FRANCHISE_RELATION_TYPES.has(edge.relationType));
}

function franchiseTargets(node: AnilistMediaNode): number[] {
  return node.relations
    .filter((edge) => FRANCHISE_RELATION_TYPES.has(edge.relationType))
    .map((edge) => edge.targetId);
}

// TV first, then other episodic formats, then films/OVAs/specials, then entries with no data.
function formatRank(format: string | null): number {
  if (format === 'TV') return 0;
  if (format === 'TV_SHORT' || format === 'ONA') return 1;
  return format === null ? 3 : 2;
}

export function buildFranchiseGraph(nodes: Iterable<AnilistMediaNode>): FranchiseGraph {
  const byId = new Map<number, AnilistMediaNode>();
  for (const node of nodes) byId.set(node.anilistId, node);

  const parent = new Map<number, number>();
  const find = (id: number): number => {
    let top = parent.get(id) ?? id;
    if (!parent.has(id)) parent.set(id, id);
    while (parent.get(top) !== top) top = parent.get(top)!;
    // Path compression keeps long sequel chains cheap.
    let at = id;
    while (at !== top) {
      const next = parent.get(at)!;
      parent.set(at, top);
      at = next;
    }
    return top;
  };

  for (const node of byId.values()) {
    find(node.anilistId);
    for (const targetId of franchiseTargets(node)) {
      const target = byId.get(targetId);
      if (target && vetoes(target, node.anilistId)) continue;
      const a = find(node.anilistId);
      const b = find(targetId);
      if (a !== b) parent.set(b, a);
    }
  }

  const members = new Map<number, number[]>();
  for (const id of parent.keys()) {
    const top = find(id);
    const list = members.get(top);
    if (list) list.push(id);
    else members.set(top, [id]);
  }

  // A franchise is named after its first TV series, else its first episodic
  // entry, else its earliest entry; ties go to the lower id.
  const compareRoots = (a: number, b: number): number => {
    const left = byId.get(a);
    const right = byId.get(b);
    return (
      formatRank(left?.format ?? null) - formatRank(right?.format ?? null) ||
      (left?.startDate ?? Number.MAX_SAFE_INTEGER) -
        (right?.startDate ?? Number.MAX_SAFE_INTEGER) ||
      a - b
    );
  };
  const roots = new Map<number, number>();
  for (const ids of members.values()) {
    const root = [...ids].sort(compareRoots)[0]!;
    for (const id of ids) roots.set(id, root);
  }

  return {
    node: (anilistId) => byId.get(anilistId),
    rootOf: (anilistId) => roots.get(anilistId),
  };
}

const FETCH_CHUNK_SIZE = 50;

export interface FranchiseWalkOptions {
  /** AniList ids the walk starts from (the library's own entries). */
  seeds: readonly number[];
  /** Cached entries; updated in place with everything fetched. */
  known: Map<number, AnilistMediaNode>;
  isFresh: (node: AnilistMediaNode) => boolean;
  /** Entries AniList returned for `ids` (at most 50). Throws on transport or API failure. */
  fetchNodes: (ids: number[]) => Promise<AnilistMediaNode[]>;
  /** Called once per fetched batch so progress survives a later failure. */
  onBatch: (nodes: AnilistMediaNode[]) => void;
  now: () => number;
  /** Hops from a seed; entries at this depth are fetched but not followed. */
  maxDepth: number;
  /** Most entries fetched in one walk; the rest wait for the next walk. */
  maxFetch: number;
  /** When true, the walk starts no further request and drops an in-flight result. */
  isCancelled?: () => boolean;
}

export interface FranchiseWalkResult {
  fetched: number;
  /** True when the fetch budget ran out before every stale entry was refreshed. */
  truncated: boolean;
}

/**
 * Breadth-first walk from the seeds along franchise relations, fetching
 * entries that are missing or stale. Fresh cached entries are followed without
 * a request. An entry AniList does not return keeps its previous data (or gets
 * an empty placeholder) so it is not asked about again until it goes stale; a
 * thrown fetch error stops the walk and leaves the cache as it was.
 */
export async function walkFranchiseRelations(
  options: FranchiseWalkOptions,
): Promise<FranchiseWalkResult> {
  const { known } = options;
  const visited = new Set<number>();
  // `from` lists every entry the walk reached this one through at this depth
  // (empty for a seed), so the order seeds are given in cannot change the result.
  let frontier: Array<{ id: number; from: number[] }> = [...new Set(options.seeds)].map((id) => ({
    id,
    from: [],
  }));
  let fetched = 0;
  let truncated = false;

  for (let depth = 0; frontier.length > 0; depth += 1) {
    for (const step of frontier) visited.add(step.id);

    const stale = frontier
      .map((step) => step.id)
      .filter((id) => {
        const node = known.get(id);
        return !node || !options.isFresh(node);
      });
    const allowed = stale.slice(0, Math.max(0, options.maxFetch - fetched));
    if (allowed.length < stale.length) truncated = true;

    for (let start = 0; start < allowed.length; start += FETCH_CHUNK_SIZE) {
      if (options.isCancelled?.()) return { fetched, truncated: true };
      const chunk = allowed.slice(start, start + FETCH_CHUNK_SIZE);
      const fetchedNodes = await options.fetchNodes(chunk);
      if (options.isCancelled?.()) return { fetched, truncated: true };
      const found = new Map(fetchedNodes.map((node) => [node.anilistId, node] as const));
      const fetchedAtMs = options.now();
      const batch = chunk.map((id): AnilistMediaNode => {
        const node = found.get(id);
        if (node) return node;
        const previous = known.get(id);
        if (previous) return { ...previous, fetchedAtMs };
        return {
          anilistId: id,
          format: null,
          titleRomaji: null,
          titleEnglish: null,
          titleNative: null,
          startDate: null,
          episodes: null,
          relations: [],
          fetchedAtMs,
        };
      });
      for (const node of batch) known.set(node.anilistId, node);
      options.onBatch(batch);
      fetched += chunk.length;
    }

    if (depth >= options.maxDepth) break;

    const next = new Map<number, number[]>();
    for (const step of frontier) {
      const node = known.get(step.id);
      if (!node) continue;
      // Reached only through links the entry itself rejects (a spin-off's
      // original): it is not part of the franchise, so its own links are not
      // followed. One accepted incoming link is enough.
      if (step.from.length > 0 && step.from.every((fromId) => vetoes(node, fromId))) continue;
      for (const targetId of franchiseTargets(node)) {
        if (visited.has(targetId)) continue;
        const from = next.get(targetId);
        if (from) from.push(node.anilistId);
        else next.set(targetId, [node.anilistId]);
      }
    }
    frontier = [...next].map(([id, from]) => ({ id, from }));
  }

  return { fetched, truncated };
}
