import { sanitizeMediaTitle } from '../../../shared/media-identity';

/**
 * AniList has no concept of "season N": sequels are separate media with their own
 * titles (Zoku, Kan, 2nd Season, ...). Searching "<title> Season 3" therefore returns
 * nothing and callers silently fall back to the season 1 entry.
 *
 * This module resolves a parsed (title, season) pair to the right media by locating a
 * season 1 anchor and then walking AniList SEQUEL relations, with an air-order fallback
 * when the relation chain is incomplete. When the seasonal entry cannot be located it
 * reports `seasonResolved: false` so callers can refuse to act instead of guessing.
 *
 * Not every SEQUEL hop is a new season. Along the chain:
 * - the back half of a split cour ("Part 2", "Cour 2", "第2クール") is its own entry,
 * - movies, specials and OVAs sit between seasons,
 * so the walk only counts entries that start a new TV season, and prefers an entry whose
 * title names the requested season ("Season 3", "3rd Season", "III", "第3期") over the
 * count. A continuation with no marker anywhere is indistinguishable from a new season;
 * the manual AniList override covers those.
 */

import { normalizeTitleIdentity } from '../../utils/title-normalization';

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

export type AnilistSeasonResolutionVia = 'anchor' | 'named-season' | 'sequel-chain' | 'air-order';

/**
 * The entry that actually airs the parsed episode when it runs past the season entry
 * into a later part of the same season ("S1E15" of an 11 + 12 split cour is episode 4
 * of the "Cour 2" entry).
 */
export interface AnilistEpisodeEntry {
  id: number;
  title: string;
  episodes: number | null;
  media: AnilistSeasonMedia;
  /** Episode number within this entry. */
  episode: number;
  exactTitleMatch: boolean;
}

export interface AnilistSeasonResolution {
  /**
   * The season's own entry (its first part). Season-level consumers (cover art, the
   * anime row's AniList link, the character dictionary) use this so every episode of a
   * split-cour season agrees on one id.
   */
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
  /** Set only when the parsed episode belongs to a later part of the season. */
  episodeEntry: AnilistEpisodeEntry | null;
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

/** An ONA this short is a special released online, not a season. */
const MIN_ONA_SEASON_EPISODES = 3;

/** Largest season the chain walk will count to. */
const MAX_SEQUEL_HOPS = 12;

/**
 * Hops the walk may take in total. Continuations and specials add hops without adding
 * seasons, so this has to leave room beyond the season number.
 */
const MAX_SEQUEL_STEPS = 24;

/** Parts a single season can be split into before the episode remap gives up. */
const MAX_SEASON_PARTS = 6;

/**
 * Markers AniList appends when one broadcast season is listed as several entries,
 * matched against normalized titles. Anchored to the end so a numbered *work*
 * ("JoJo no Kimyou na Bouken Part 5: Ougon no Kaze") is not read as a continuation:
 * those carry a subtitle after the number and really are separate seasons.
 */
const PART_MARKER_PATTERNS: RegExp[] = [
  /\bpart\s*(\d{1,2}|i{1,3}|iv|v)$/u,
  /\bcour\s*(\d{1,2})$/u,
  /\b(\d{1,2})(?:st|nd|rd|th)\s+(?:part|cour)$/u,
  /\b(second|final)\s+(?:part|cour)$/u,
  /第\s*(\d{1,2})\s*(?:部|クール)$/u,
];

const NUMBER_WORDS: Record<string, number> = {
  i: 1,
  ii: 2,
  iii: 3,
  iv: 4,
  v: 5,
  vi: 6,
  second: 2,
  third: 3,
  fourth: 4,
  fifth: 5,
  sixth: 6,
  一: 1,
  二: 2,
  三: 3,
  四: 4,
  五: 5,
  六: 6,
  七: 7,
  八: 8,
  九: 9,
};

/**
 * Patterns that state a season outright. Safe on any title or synonym, because none of
 * them can be confused with a part or cour number.
 */
const EXPLICIT_SEASON_PATTERNS: RegExp[] = [
  /\bseason\s*(\d{1,2})\b/u,
  /\b(\d{1,2})(?:st|nd|rd|th)\s+season\b/u,
  /\b(second|third|fourth|fifth|sixth)\s+season\b/u,
  /\bs(\d{1,2})\b/u,
  /第\s*(\d{1,2}|[二三四五六七八九])\s*[期季]/u,
  /(\d{1,2})\s*(?:期|기)/u,
];

/** Words after which a trailing number counts parts, not seasons. */
const PART_WORDS = new Set(['part', 'cour', 'parte', 'partie']);

function parseNumberToken(token: string): number | null {
  if (/^\d{1,2}$/.test(token)) return Number(token);
  return NUMBER_WORDS[token] ?? null;
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

/**
 * Splits a trailing part/cour marker off a title, or returns null when it has none.
 *
 * "Mushoku Tensei: Jobless Reincarnation Season 2 Part 2" is the back half of season 2,
 * not season 3. `part` is null for an unnumbered marker ("Final Cour").
 */
export function splitPartMarker(title: string): { base: string; part: number | null } | null {
  const normalized = normalizeTitleIdentity(title);
  for (const pattern of PART_MARKER_PATTERNS) {
    const match = normalized.match(pattern);
    if (!match || match.index === undefined) continue;
    const base = normalized.slice(0, match.index).trim();
    if (!base) continue;
    const token = match[1] ?? '';
    return { base, part: token === 'final' ? null : parseNumberToken(token) };
  }
  return null;
}

function mainTitles(media: AnilistSeasonMedia): string[] {
  return [media.title?.english, media.title?.romaji, media.title?.native]
    .filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
    .map((value) => normalizeTitleIdentity(value));
}

function mediaTitles(media: AnilistSeasonMedia): string[] {
  const synonyms = (Array.isArray(media.synonyms) ? media.synonyms : [])
    .filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
    .map((value) => normalizeTitleIdentity(value));
  return [...mainTitles(media), ...synonyms].filter((value) => value.length > 0);
}

/**
 * Season numbers an entry's titles name.
 *
 * A bare trailing number ("Overlord II", "Konosuba 2") is only trusted on the main
 * titles: localized synonyms end in part numbers spelled in other scripts (Mushoku
 * Tensei's "Cour 2" carries the Thai "... พาร์ท 2", "part 2").
 */
export function namedSeasons(media: AnilistSeasonMedia): Set<number> {
  const seasons = new Set<number>();
  for (const title of mediaTitles(media)) {
    for (const pattern of EXPLICIT_SEASON_PATTERNS) {
      const token = title.match(pattern)?.[1];
      const value = token ? parseNumberToken(token) : null;
      if (value !== null && value > 0) seasons.add(value);
    }
  }
  for (const title of mainTitles(media)) {
    const words = title.split(' ');
    const last = words.at(-1) ?? '';
    const before = words.at(-2) ?? '';
    if (PART_WORDS.has(before)) continue;
    const value = parseNumberToken(last);
    if (value !== null && value > 1) seasons.add(value);
  }
  return seasons;
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

/** Whether an entry can be a season: a TV series, or an ONA long enough not to be a special. */
function isSeasonEntry(media: AnilistSeasonMedia): boolean {
  if (!isSeasonalFormat(media)) return false;
  if ((media.format || '').toUpperCase() !== 'ONA') return true;
  const episodes = episodeCount(media);
  return episodes === null || episodes >= MIN_ONA_SEASON_EPISODES;
}

function formatRank(media: AnilistSeasonMedia): number {
  if (!isSeasonEntry(media)) return SEASONAL_FORMAT_PRIORITY.length;
  return SEASONAL_FORMAT_PRIORITY.indexOf((media.format || '').toUpperCase());
}

/**
 * Whether `next` continues the season `seasonStart` began rather than starting a new one.
 *
 * Compared against the season's first entry, not the previous hop, so a special between
 * a season and its second cour cannot hide the link. The marker is required, not just a
 * matching base: a sequel that repeats its predecessor's name exactly is a new season
 * AniList did not number. "Part 1" never marks a back half, and a season's own marker is
 * only stripped when it is "Part 1" (AoT season 3 carries "... ภาค 3 Part 1"); stripping
 * any number would read JoJo's "Part 6" as a continuation of its "Part 5".
 */
function isSeasonContinuation(seasonStart: AnilistSeasonMedia, next: AnilistSeasonMedia): boolean {
  const markers = mediaTitles(next)
    .map((title) => splitPartMarker(title))
    .filter((marker): marker is NonNullable<typeof marker> => marker !== null && marker.part !== 1);
  if (markers.length === 0) return false;

  const seasonKeys = new Set<string>();
  for (const title of mediaTitles(seasonStart)) {
    seasonKeys.add(title);
    const own = splitPartMarker(title);
    if (own?.part === 1) seasonKeys.add(own.base);
  }
  return markers.some((marker) => seasonKeys.has(marker.base));
}

function exactMatchFor(candidate: AnilistSeasonMedia, titles: string[]): boolean {
  const candidateTitles = mediaTitles(candidate);
  return titles.some((title) => candidateTitles.includes(normalizeTitleIdentity(title)));
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

  const targets = [
    normalizeTitleIdentity(title),
    normalizeTitleIdentity(stripSeasonSuffix(title)),
  ].filter((value, index, all) => value.length > 0 && all.indexOf(value) === index);

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

/** SEQUEL entries of a media id, minus those already visited, best continuation first. */
type SequelLookup = (id: number, visited: Set<number>) => Promise<AnilistSeasonMedia[]>;

/**
 * Relation lookups for one resolution, memoized so the episode remap does not refetch
 * what the chain walk already loaded. Sorted with a season over a special, then the TV
 * line over other formats, then the earliest. Transport errors propagate: a failed lookup
 * must not be mistaken for "no sequel exists".
 */
function createSequelLookup(deps: ResolveAnilistSeasonMediaDeps): SequelLookup {
  const cache = new Map<number, Promise<AnilistSeasonMedia[]>>();
  const load = async (id: number): Promise<AnilistSeasonMedia[]> => {
    const response = await deps.execute<AnilistSeasonRelationsResponse>(
      ANILIST_SEASON_RELATIONS_QUERY,
      { id },
    );
    const sequels: AnilistSeasonMedia[] = [];
    for (const edge of response.Media?.relations?.edges ?? []) {
      const node = edge?.node;
      if (!node || typeof node.id !== 'number') continue;
      if ((edge.relationType || '').toUpperCase() !== 'SEQUEL') continue;
      if ((node.type || 'ANIME').toUpperCase() !== 'ANIME') continue;
      sequels.push(node);
    }
    return sequels.sort((a, b) => {
      const rankDelta = formatRank(a) - formatRank(b);
      if (rankDelta !== 0) return rankDelta;
      const yearDelta =
        (airYear(a) ?? Number.MAX_SAFE_INTEGER) - (airYear(b) ?? Number.MAX_SAFE_INTEGER);
      if (yearDelta !== 0) return yearDelta;
      return a.id - b.id;
    });
  };
  return async (id, visited) => {
    let pending = cache.get(id);
    if (!pending) {
      pending = load(id);
      cache.set(id, pending);
      // A failed lookup is retried by the next caller rather than cached.
      pending.catch(() => cache.delete(id));
    }
    return (await pending).filter((node) => !visited.has(node.id));
  };
}

type ChainResult = { media: AnilistSeasonMedia; via: 'named-season' | 'sequel-chain' };

async function walkSequelChain(
  anchor: AnilistSeasonMedia,
  season: number,
  sequelsOf: SequelLookup,
): Promise<ChainResult | null> {
  // Advancing fewer seasons than requested would land on the wrong one and report it as
  // resolved, so refuse instead and let the caller fall through to its guarded fallback.
  const targetAdvances = season - 1;
  if (targetAdvances > MAX_SEQUEL_HOPS) {
    return null;
  }
  const visited = new Set<number>([anchor.id]);
  let current = anchor;
  // The last entry that counted as a season.
  let seasonStart = anchor;
  let advances = 0;
  // The unnamed entry the count landed on. The walk looks one season further before
  // accepting it, because an entry naming the requested season outranks the count: the
  // counted entry may be a continuation with no marker.
  let counted: AnilistSeasonMedia | null = null;
  // Set once counting reached the requested season on an entry whose title names a
  // different one. Only a named match can answer after that.
  let countContradicted = false;

  for (let step = 0; step < MAX_SEQUEL_STEPS; step += 1) {
    const next = (await sequelsOf(current.id, visited))[0];
    if (!next) break;
    current = next;
    visited.add(next.id);

    // Specials, movies and OVAs are walked through, because the rest of the franchise
    // hangs off them, but never counted.
    if (!isSeasonEntry(next)) continue;

    const named = namedSeasons(next);
    if (isSeasonContinuation(seasonStart, next)) {
      // A back half naming the requested season ("Season 3 Part 2") vouches for the
      // entry its season started on.
      if (named.has(season)) return { media: seasonStart, via: 'named-season' };
      continue;
    }
    if (named.has(season)) return { media: next, via: 'named-season' };
    // The season after the counted one does not claim the requested number either.
    if (counted) break;
    if (countContradicted) {
      // Past the requested season by name: it is not on this chain.
      if ([...named].some((value) => value > season)) return null;
      continue;
    }

    advances += 1;
    seasonStart = next;
    if (advances === targetAdvances) {
      if (named.size === 0) counted = next;
      else countContradicted = true;
    }
  }

  return counted ? { media: counted, via: 'sequel-chain' } : null;
}

/**
 * Whether the parsed episode runs past `seasonEntry` into a later part of the same season,
 * and if so, which entry airs it and at what episode number. Null when the episode fits
 * the season entry, its length is unknown, or no continuation covers the episode (e.g.
 * absolute numbering across seasons).
 */
async function findEpisodeEntry(
  seasonEntry: AnilistSeasonMedia,
  episode: number,
  sequelsOf: SequelLookup,
): Promise<{ media: AnilistSeasonMedia; episode: number } | null> {
  const seasonLength = episodeCount(seasonEntry);
  if (seasonLength === null || episode <= seasonLength) return null;

  const visited = new Set<number>([seasonEntry.id]);
  let current = seasonEntry;
  // The last part whose episodes are still counted in `remaining`; its length is known.
  let lastPart = seasonEntry;
  let remaining = episode;
  let parts = 1;

  for (let step = 0; step < MAX_SEQUEL_STEPS; step += 1) {
    const next = (await sequelsOf(current.id, visited))[0];
    if (!next) return null;
    visited.add(next.id);
    current = next;
    // A recap special between the two halves airs none of the season's episodes.
    if (!isSeasonEntry(next)) continue;
    if (!isSeasonContinuation(seasonEntry, next)) return null;

    remaining -= episodeCount(lastPart) ?? 0;
    lastPart = next;
    parts += 1;
    const length = episodeCount(next);
    if (length === null || remaining <= length) return { media: next, episode: remaining };
    if (parts >= MAX_SEASON_PARTS) return null;
  }
  return null;
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
  // Any of the anchor's main titles, matched on word boundaries. Synonyms are too short
  // and loose ("SnK", "ReZero") to define a franchise, and a single title misses entries
  // that only spell another one out (Mushoku Tensei II's native title adds "Ⅱ" to the
  // word the season 1 native title ends on).
  const anchorBases = mainTitles(anchor)
    .map((value) => stripSeasonSuffix(value))
    .filter((value) => value.length > 0);
  if (anchorBases.length === 0) return null;

  const franchise = media.filter((entry) => {
    if (!isSeasonEntry(entry)) return false;
    if (entry.id === anchor.id) return true;
    return mediaTitles(entry).some((candidate) => {
      const padded = ` ${stripSeasonSuffix(candidate)} `;
      return anchorBases.some((base) => padded.includes(` ${base} `));
    });
  });

  // Drop the back halves of split-cour seasons, which are separate entries here but not
  // separate seasons: counting them shifts every later season index. An entry is only
  // compared against the *other* entries' titles, so a season whose own localized synonym
  // spells it "<season> Part 1" is kept.
  const seasons = franchise.filter((entry) => {
    if (entry.id === anchor.id) return true;
    return !franchise.some((other) => other.id !== entry.id && isSeasonContinuation(other, entry));
  });
  if (seasons.length < season) return null;

  // Ordering by air date is only meaningful when every entry has one: an unknown year
  // would sort last and shift the season index while still reporting a resolved match.
  if (seasons.some((entry) => airYear(entry) === null)) return null;

  const ordered = [...seasons].sort((a, b) => {
    const yearDelta = (airYear(a) ?? 0) - (airYear(b) ?? 0);
    if (yearDelta !== 0) return yearDelta;
    return a.id - b.id;
  });

  // The anchor has to be the first entry, otherwise this ordering is not a season list.
  if (ordered[0]?.id !== anchor.id) return null;

  const picked = ordered[season - 1] ?? null;
  if (!picked) return null;
  // An entry that names some other season shows this ordering is off.
  const named = namedSeasons(picked);
  if (named.size > 0 && !named.has(season)) return null;
  return picked;
}

export async function resolveAnilistSeasonMedia(
  input: ResolveAnilistSeasonMediaInput,
  deps: ResolveAnilistSeasonMediaDeps,
): Promise<AnilistSeasonResolution | null> {
  const safeTitle = sanitizeMediaTitle(input.title);
  if (!safeTitle) return null;
  const searchTitle = stripSeasonSuffix(safeTitle).trim() || safeTitle;
  if (!searchTitle) return null;

  const season =
    typeof input.season === 'number' && Number.isInteger(input.season) && input.season > 0
      ? input.season
      : null;
  const episode =
    typeof input.episode === 'number' && Number.isInteger(input.episode) && input.episode > 0
      ? input.episode
      : null;

  const response = await deps.execute<AnilistSeasonSearchResponse>(ANILIST_SEASON_SEARCH_QUERY, {
    search: searchTitle,
  });
  const media = (response.Page?.media ?? []).filter(
    (entry): entry is AnilistSeasonMedia => Boolean(entry) && typeof entry.id === 'number',
  );
  if (media.length === 0) return null;

  // Certifies the media actually returned, never the anchor on its behalf: a hop can land
  // on the wrong entry while the anchor title still matches perfectly, and that certainty
  // must not carry over to the hopped-to entry.
  const parsedTitles = [searchTitle, input.title];
  const sequelsOf = createSequelLookup(deps);

  const finish = (
    seasonEntry: AnilistSeasonMedia,
    via: AnilistSeasonResolutionVia,
    split: { media: AnilistSeasonMedia; episode: number } | null = null,
    seasonResolved = true,
  ): AnilistSeasonResolution => {
    if (split) {
      deps.logInfo?.(
        `[anilist] episode ${episode} of "${searchTitle}" is episode ${split.episode} of ${displayTitle(split.media, searchTitle)} (${split.media.id})`,
      );
    }
    return {
      id: seasonEntry.id,
      title: displayTitle(seasonEntry, searchTitle),
      episodes: episodeCount(seasonEntry),
      media: seasonEntry,
      seasonResolved,
      requestedSeason: season,
      via,
      exactTitleMatch: exactMatchFor(seasonEntry, parsedTitles),
      episodeEntry: split
        ? {
            id: split.media.id,
            title: displayTitle(split.media, searchTitle),
            episodes: episodeCount(split.media),
            media: split.media,
            episode: split.episode,
            exactTitleMatch: exactMatchFor(split.media, parsedTitles),
          }
        : null,
    };
  };

  if (season === null || season <= 1) {
    // The best title match first. If the episode runs past it, a split-cour continuation
    // may air it ("S1E15" of 11 + 12); only when none does is the episode count used to
    // pick a different, longer entry (a TV series over a same-titled movie).
    const titleAnchor = pickAnchorMedia(searchTitle, media);
    if (!titleAnchor) return null;
    if (episode === null) return finish(titleAnchor, 'anchor');
    const split = await findEpisodeEntry(titleAnchor, episode, sequelsOf);
    if (split) return finish(titleAnchor, 'anchor', split);
    const fitted = pickAnchorMedia(searchTitle, media, { episode }) ?? titleAnchor;
    return finish(fitted, 'anchor');
  }

  // Later seasons must not filter by episode, because the anchor is season 1 and may be
  // shorter than the requested episode.
  const anchor = pickAnchorMedia(searchTitle, media);
  if (!anchor) return null;

  // Transport failures propagate from here: the air-order fallback is for a chain that
  // answered and had no edge to follow, and running it after a failed hop would turn a
  // rate limit into a confident wrong season.
  const viaChain = await walkSequelChain(anchor, season, sequelsOf);
  const seasonEntry = viaChain?.media ?? pickByAirOrder(anchor, season, media);
  if (seasonEntry) {
    const via = viaChain?.via ?? 'air-order';
    deps.logInfo?.(
      `[anilist] season ${season} of "${searchTitle}" resolved via ${via}: ${displayTitle(seasonEntry, searchTitle)} (${seasonEntry.id})`,
    );
    const split = episode === null ? null : await findEpisodeEntry(seasonEntry, episode, sequelsOf);
    return finish(seasonEntry, via, split);
  }

  deps.logInfo?.(
    `[anilist] could not resolve season ${season} of "${searchTitle}"; falling back to ${displayTitle(anchor, searchTitle)} (${anchor.id})`,
  );
  return finish(anchor, 'anchor', null, false);
}
