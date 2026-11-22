import { executeAnilistQuery } from './anilist-graphql';
import type { AnilistMediaNode, AnilistRelationEdge } from './franchise-graph';
import type { AnilistRateLimiter } from './rate-limiter';

/** Loads AniList entries with their relations, for the library's franchise grouping. */
export interface AnilistRelationFetcher {
  /** Entries AniList returned for `ids` (at most 50); ids it does not know are absent. */
  fetchMedia(ids: number[]): Promise<AnilistMediaNode[]>;
}

const RELATIONS_QUERY = `
query ($ids: [Int], $perPage: Int) {
  Page(page: 1, perPage: $perPage) {
    media(id_in: $ids, type: ANIME) {
      id
      format
      episodes
      title { romaji english native }
      startDate { year month day }
      relations { edges { relationType(version: 2) node { id type } } }
    }
  }
}`;

interface RelationsMedia {
  id: number;
  format: string | null;
  episodes: number | null;
  title: { romaji: string | null; english: string | null; native: string | null } | null;
  startDate: { year: number | null; month: number | null; day: number | null } | null;
  relations: {
    edges: Array<{
      relationType: string | null;
      node: { id: number; type: string | null } | null;
    } | null> | null;
  } | null;
}

interface RelationsResponse {
  Page: { media: Array<RelationsMedia | null> | null } | null;
}

export function toAnilistMediaNode(media: RelationsMedia, fetchedAtMs: number): AnilistMediaNode {
  const start = media.startDate;
  const relations: AnilistRelationEdge[] = [];
  for (const edge of media.relations?.edges ?? []) {
    // Manga and novels are not part of a watch library.
    if (!edge?.relationType || edge.node?.type !== 'ANIME') continue;
    relations.push({ relationType: edge.relationType, targetId: edge.node.id });
  }
  return {
    anilistId: media.id,
    format: media.format,
    titleRomaji: media.title?.romaji ?? null,
    titleEnglish: media.title?.english ?? null,
    titleNative: media.title?.native ?? null,
    startDate: start?.year
      ? start.year * 10_000 + (start.month ?? 0) * 100 + (start.day ?? 0)
      : null,
    episodes: media.episodes,
    relations,
    fetchedAtMs,
  };
}

export function createAnilistRelationFetcher(
  rateLimiter: AnilistRateLimiter,
  now: () => number = Date.now,
): AnilistRelationFetcher {
  return {
    async fetchMedia(ids) {
      if (ids.length === 0) return [];
      const data = await executeAnilistQuery<RelationsResponse>(rateLimiter, RELATIONS_QUERY, {
        ids,
        perPage: ids.length,
      });
      const fetchedAtMs = now();
      return (data.Page?.media ?? [])
        .filter((media): media is RelationsMedia => media !== null)
        .map((media) => toAnilistMediaNode(media, fetchedAtMs));
    },
  };
}
