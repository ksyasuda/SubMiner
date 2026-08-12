/**
 * AniList has no concept of "season N": sequels are separate media with their own
 * titles (Zoku, Kan, 2nd Season, ...). Searching "<title> Season 3" therefore returns
 * nothing and callers silently fall back to the season 1 entry.
 *
 * This module resolves a parsed (title, season) pair to the right media by locating a
 * season 1 anchor and then walking AniList SEQUEL relations, with an air-order fallback
 * when the relation chain is incomplete. When the seasonal entry cannot be located it
 * reports `seasonResolved: false` so callers can refuse to act instead of guessing.
 */

export interface AnilistSeasonMediaTitle {
  romaji?: string | null;
  english?: string | null;
  native?: string | null;
}

export interface AnilistSeasonMedia {
  id: number;
  episodes?: number | null;
  format?: string | null;
  seasonYear?: number | null;
  startDate?: { year?: number | null } | null;
  synonyms?: Array<string | null> | null;
  coverImage?: { large?: string | null; medium?: string | null } | null;
  title?: AnilistSeasonMediaTitle | null;
}

export type AnilistQueryExecutor = <T>(
  query: string,
  variables: Record<string, unknown>,
) => Promise<T>;

export type AnilistSeasonResolutionVia = 'anchor' | 'sequel-chain' | 'air-order';

export interface AnilistSeasonResolution {
  id: number;
  title: string;
  episodes: number | null;
  media: AnilistSeasonMedia;
  /** False when a season >= 2 was requested but no seasonal entry could be located. */
  seasonResolved: boolean;
  requestedSeason: number | null;
  via: AnilistSeasonResolutionVia;
  /** Exact normalized match against an AniList title or synonym. */
  exactTitleMatch: boolean;
}

export interface ResolveAnilistSeasonMediaInput {
  title: string;
  season?: number | null;
  episode?: number | null;
}

export interface ResolveAnilistSeasonMediaDeps {
  execute: AnilistQueryExecutor;
  logInfo?: (message: string) => void;
}

const MEDIA_FIELDS = `
  id
  episodes
  format
  seasonYear
  startDate { year }
  synonyms
  coverImage { large medium }
  title { romaji english native }
`;

export const ANILIST_SEASON_SEARCH_QUERY = `
query ($search: String!) {
  Page(perPage: 10) {
    media(search: $search, type: ANIME, sort: [SEARCH_MATCH, POPULARITY_DESC]) {
      ${MEDIA_FIELDS}
    }
  }
}
`;

export const ANILIST_SEASON_RELATIONS_QUERY = `
query ($id: Int!) {
  Media(id: $id, type: ANIME) {
    id
    relations {
      edges {
        relationType
        node {
          type
          ${MEDIA_FIELDS}
        }
      }
    }
  }
}
`;

interface AnilistSeasonSearchResponse {
  Page?: {
    media?: AnilistSeasonMedia[] | null;
  } | null;
}

interface AnilistSeasonRelationsResponse {
  Media?: {
    relations?: {
      edges?: Array<{
        relationType?: string | null;
        node?: (AnilistSeasonMedia & { type?: string | null }) | null;
      } | null> | null;
    } | null;
  } | null;
}

/** Formats that can carry a numbered TV season, best first. */
const SEASONAL_FORMAT_PRIORITY = ['TV', 'TV_SHORT', 'ONA'];

const MAX_SEQUEL_HOPS = 12;

function normalizeTitle(value: string): string {
  return value
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

/**
 * Drops season markers a release name carries but AniList titles never do,
 * so "Some Show Season 3" and "Some Show S3" both search as "Some Show".
 */
export function stripSeasonSuffix(title: string): string {
  return title
    .replace(/\bseason\s*\d{1,2}\b/gi, ' ')
    .replace(/\b\d{1,2}(?:st|nd|rd|th)\s+season\b/gi, ' ')
    .replace(/\bs\d{1,2}\b/gi, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function mediaTitles(media: AnilistSeasonMedia): string[] {
  const synonyms = Array.isArray(media.synonyms) ? media.synonyms : [];
  return [media.title?.english, media.title?.romaji, media.title?.native, ...synonyms]
    .filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
    .map((value) => normalizeTitle(value));
}

function displayTitle(media: AnilistSeasonMedia, fallback: string): string {
  return (
    media.title?.english?.trim() ||
    media.title?.romaji?.trim() ||
    media.title?.native?.trim() ||
    fallback.trim()
  );
}

function episodeCount(media: AnilistSeasonMedia): number | null {
  return typeof media.episodes === 'number' && media.episodes > 0 ? media.episodes : null;
}

function airYear(media: AnilistSeasonMedia): number | null {
  if (typeof media.seasonYear === 'number' && media.seasonYear > 0) {
    return media.seasonYear;
  }
  const startYear = media.startDate?.year;
  return typeof startYear === 'number' && startYear > 0 ? startYear : null;
}

function isSeasonalFormat(media: AnilistSeasonMedia): boolean {
  const format = (media.format || '').toUpperCase();
  return SEASONAL_FORMAT_PRIORITY.includes(format);
}

function formatRank(media: AnilistSeasonMedia): number {
  const index = SEASONAL_FORMAT_PRIORITY.indexOf((media.format || '').toUpperCase());
  return index < 0 ? SEASONAL_FORMAT_PRIORITY.length : index;
}

function toResolution(
  media: AnilistSeasonMedia,
  fallbackTitle: string,
  season: number | null,
  via: AnilistSeasonResolutionVia,
  seasonResolved: boolean,
  exactTitleMatch: boolean,
): AnilistSeasonResolution {
  return {
    id: media.id,
    title: displayTitle(media, fallbackTitle),
    episodes: episodeCount(media),
    media,
    seasonResolved,
    requestedSeason: season,
    via,
    exactTitleMatch,
  };
}

/**
 * Picks the franchise anchor (season 1 entry) for a search result set. The anchor is
 * matched on title alone - season is handled by walking relations from here.
 */
export function pickAnchorMedia(
  title: string,
  media: AnilistSeasonMedia[],
  options: { episode?: number | null } = {},
): AnilistSeasonMedia | null {
  if (media.length === 0) return null;

  const episode = options.episode;
  const episodeFiltered =
    typeof episode === 'number' && episode > 0
      ? media.filter((entry) => {
          const total = episodeCount(entry);
          return total === null || total >= episode;
        })
      : media;
  const pool = episodeFiltered.length > 0 ? episodeFiltered : media;

  const targets = [normalizeTitle(title), normalizeTitle(stripSeasonSuffix(title))].filter(
    (value, index, all) => value.length > 0 && all.indexOf(value) === index,
  );

  const scored = pool.map((entry, index) => {
    const candidateTitles = mediaTitles(entry);
    let score = 0;
    for (const target of targets) {
      if (candidateTitles.includes(target)) {
        score += 120;
        continue;
      }
      if (candidateTitles.some((candidate) => candidate.startsWith(target))) {
        score += 40;
      } else if (candidateTitles.some((candidate) => candidate.includes(target))) {
        score += 25;
      }
      if (candidateTitles.some((candidate) => target.includes(candidate))) {
        score += 10;
      }
    }
    if (isSeasonalFormat(entry)) {
      score += 30;
    }
    return { entry, score, index };
  });

  scored.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    if (a.index !== b.index) return a.index - b.index;
    return a.entry.id - b.entry.id;
  });

  return scored[0]?.entry ?? null;
}

async function walkSequelChain(
  anchor: AnilistSeasonMedia,
  season: number,
  deps: ResolveAnilistSeasonMediaDeps,
): Promise<AnilistSeasonMedia | null> {
  // Walking fewer hops than requested would land on the wrong season and report it as
  // resolved, so refuse instead and let the caller fall through to its guarded fallback.
  const hops = season - 1;
  if (hops > MAX_SEQUEL_HOPS) {
    return null;
  }
  const visited = new Set<number>([anchor.id]);
  let current = anchor;

  for (let hop = 0; hop < hops; hop += 1) {
    // Transport errors propagate: a failed hop must not be mistaken for "no sequel exists".
    const response = await deps.execute<AnilistSeasonRelationsResponse>(
      ANILIST_SEASON_RELATIONS_QUERY,
      { id: current.id },
    );

    const sequels = (response.Media?.relations?.edges ?? [])
      .map((edge) => edge ?? null)
      .filter(
        (
          edge,
        ): edge is {
          relationType?: string | null;
          node: AnilistSeasonMedia & { type?: string | null };
        } =>
          Boolean(edge?.node) &&
          (edge?.relationType || '').toUpperCase() === 'SEQUEL' &&
          (edge?.node?.type || 'ANIME').toUpperCase() === 'ANIME',
      )
      .map((edge) => edge.node)
      .filter((node) => !visited.has(node.id));

    if (sequels.length === 0) {
      return null;
    }

    // A franchise can branch (a TV sequel plus an ONA spinoff); prefer the TV line.
    sequels.sort((a, b) => {
      const rankDelta = formatRank(a) - formatRank(b);
      if (rankDelta !== 0) return rankDelta;
      const yearDelta =
        (airYear(a) ?? Number.MAX_SAFE_INTEGER) - (airYear(b) ?? Number.MAX_SAFE_INTEGER);
      if (yearDelta !== 0) return yearDelta;
      return a.id - b.id;
    });

    current = sequels[0]!;
    visited.add(current.id);
  }

  return current.id === anchor.id ? null : current;
}

/**
 * Fallback for franchises whose SEQUEL edges are missing or route through a format we
 * skipped: order the franchise's seasonal entries by air date and index by season.
 */
export function pickByAirOrder(
  anchor: AnilistSeasonMedia,
  season: number,
  media: AnilistSeasonMedia[],
): AnilistSeasonMedia | null {
  const anchorTitles = mediaTitles(anchor).map((value) => stripSeasonSuffix(value));
  const anchorBase = anchorTitles.sort((a, b) => a.length - b.length)[0];
  if (!anchorBase) return null;

  const franchise = media.filter((entry) => {
    if (!isSeasonalFormat(entry)) return false;
    if (entry.id === anchor.id) return true;
    return mediaTitles(entry).some((candidate) =>
      stripSeasonSuffix(candidate).includes(anchorBase),
    );
  });
  if (franchise.length < season) return null;

  // Ordering by air date is only meaningful when every entry has one: an unknown year
  // would sort last and shift the season index while still reporting a resolved match.
  if (franchise.some((entry) => airYear(entry) === null)) return null;

  const ordered = [...franchise].sort((a, b) => {
    const yearDelta = (airYear(a) ?? 0) - (airYear(b) ?? 0);
    if (yearDelta !== 0) return yearDelta;
    return a.id - b.id;
  });

  // The anchor has to be the first entry, otherwise this ordering is not a season list.
  if (ordered[0]?.id !== anchor.id) return null;

  return ordered[season - 1] ?? null;
}

export async function resolveAnilistSeasonMedia(
  input: ResolveAnilistSeasonMediaInput,
  deps: ResolveAnilistSeasonMediaDeps,
): Promise<AnilistSeasonResolution | null> {
  const searchTitle = stripSeasonSuffix(input.title).trim() || input.title.trim();
  if (!searchTitle) return null;

  const season =
    typeof input.season === 'number' && Number.isInteger(input.season) && input.season > 0
      ? input.season
      : null;

  const response = await deps.execute<AnilistSeasonSearchResponse>(ANILIST_SEASON_SEARCH_QUERY, {
    search: searchTitle,
  });
  const media = (response.Page?.media ?? []).filter(
    (entry): entry is AnilistSeasonMedia => Boolean(entry) && typeof entry.id === 'number',
  );
  if (media.length === 0) return null;

  // Season 1 (or unknown) resolves against the episode count; later seasons must not,
  // because the anchor is season 1 and may be shorter than the requested episode.
  const anchor = pickAnchorMedia(searchTitle, media, {
    episode: season === null || season <= 1 ? input.episode : null,
  });
  if (!anchor) return null;
  const exactTitleMatch = mediaTitles(anchor).includes(normalizeTitle(searchTitle));

  if (season === null || season <= 1) {
    return toResolution(anchor, searchTitle, season, 'anchor', true, exactTitleMatch);
  }

  let chainError: unknown = null;
  let viaChain: AnilistSeasonMedia | null = null;
  try {
    viaChain = await walkSequelChain(anchor, season, deps);
  } catch (error) {
    chainError = error;
  }
  if (viaChain) {
    deps.logInfo?.(
      `[anilist] season ${season} of "${searchTitle}" resolved via sequel chain: ${displayTitle(viaChain, searchTitle)} (${viaChain.id})`,
    );
    return toResolution(viaChain, searchTitle, season, 'sequel-chain', true, exactTitleMatch);
  }

  const viaAirOrder = pickByAirOrder(anchor, season, media);
  if (viaAirOrder) {
    deps.logInfo?.(
      `[anilist] season ${season} of "${searchTitle}" resolved via air order: ${displayTitle(viaAirOrder, searchTitle)} (${viaAirOrder.id})`,
    );
    return toResolution(viaAirOrder, searchTitle, season, 'air-order', true, exactTitleMatch);
  }

  // The chain failed for transport reasons rather than because the season is absent;
  // surface that so callers retry instead of reporting an unresolvable season.
  if (chainError) {
    throw chainError;
  }

  deps.logInfo?.(
    `[anilist] could not resolve season ${season} of "${searchTitle}"; falling back to ${displayTitle(anchor, searchTitle)} (${anchor.id})`,
  );
  return toResolution(anchor, searchTitle, season, 'anchor', false, exactTitleMatch);
}
