import type { AnilistRateLimiter } from './rate-limiter';
import type { DatabaseSync } from '../immersion-tracker/sqlite';
import {
  getAnimeCoverArt,
  getCoverArt,
  upsertCoverArt,
  updateAnimeAnilistInfo,
} from '../immersion-tracker/query';
import {
  guessAnilistMediaInfo,
  runGuessit,
  type GuessAnilistMediaInfoDeps,
} from './anilist-updater';

const ANILIST_GRAPHQL_URL = 'https://graphql.anilist.co';
const NO_MATCH_RETRY_MS = 5 * 60 * 1000;

const SEARCH_QUERY = `
query ($search: String!) {
  Page(perPage: 5) {
    media(search: $search, type: ANIME) {
      id
      episodes
      season
      seasonYear
      coverImage { large medium }
      title { romaji english native }
    }
  }
}
`;

interface AnilistMedia {
  id: number;
  episodes: number | null;
  season: string | null;
  seasonYear: number | null;
  coverImage: { large: string | null; medium: string | null } | null;
  title: { romaji: string | null; english: string | null; native: string | null } | null;
}

interface AnilistSearchResponse {
  data?: {
    Page?: {
      media?: AnilistMedia[];
    };
  };
  errors?: Array<{ message?: string }>;
}

export interface CoverArtFetcher {
  fetchIfMissing(db: DatabaseSync, videoId: number, canonicalTitle: string): Promise<boolean>;
}

interface Logger {
  info(msg: string, ...args: unknown[]): void;
  warn(msg: string, ...args: unknown[]): void;
  error(msg: string, ...args: unknown[]): void;
}

interface CoverArtCandidate {
  title: string;
  source: 'guessit' | 'fallback';
  season: number | null;
  episode: number | null;
}

interface CoverArtFetcherOptions {
  runGuessit?: GuessAnilistMediaInfoDeps['runGuessit'];
}

export function stripFilenameTags(raw: string): string {
  let title = raw.replace(/\.[A-Za-z0-9]{2,4}$/, '');

  title = title.replace(/^(?:\s*\[[^\]]*\]\s*)+/, '');
  title = title.replace(/[._]+/g, ' ');

  // Remove everything from " - S##E##" or " - ###" onward (season/episode markers)
  title = title.replace(/\s+-\s+S\d+E\d+.*$/i, '');
  title = title.replace(/\s+-\s+\d{2,}(\s+-\s+\d+)?(\s+-.+)?$/, '');
  title = title.replace(/\s+S\d+E\d+.*$/i, '');
  title = title.replace(/\s+S\d+\s*[- ]\s*\d+[: -].*$/i, '');
  title = title.replace(/\s+E\d+[: -].*$/i, '');
  title = title.replace(/^S\d+E\d+\s*[- ]\s*/i, '');

  // Remove bracketed/parenthesized tags: [WEBDL-1080p], (2022), etc.
  title = title.replace(/\s*\[[^\]]*\]\s*/g, ' ');
  title = title.replace(/\s*\([^)]*\d{4}[^)]*\)\s*/g, ' ');

  // Remove common codec/source tags that may appear without brackets
  title = title.replace(
    /\b(WEBDL|WEBRip|BluRay|BDRip|HDTV|DVDRip|x264|x265|H\.?264|H\.?265|AV1|AAC|FLAC|Opus|10bit|8bit|1080p|720p|480p|2160p|4K)\b[-.\w]*/gi,
    '',
  );

  // Remove trailing dashes and group tags like "-Retr0"
  title = title.replace(/\s*-\s*[\w]+$/, '');

  return title.trim().replace(/\s{2,}/g, ' ');
}

function removeSeasonHint(title: string): string {
  return title
    .replace(/\bseason\s*\d+\b/gi, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function normalizeTitle(text: string): string {
  return text.trim().toLowerCase().replace(/\s+/g, ' ');
}

function extractCandidateSeasonHints(text: string): Set<number> {
  const normalized = normalizeTitle(text);
  const matches = [
    ...normalized.matchAll(/\bseason\s*(\d{1,2})\b/gi),
    ...normalized.matchAll(/\bs(\d{1,2})(?:\b|\D)/gi),
  ];
  const values = new Set<number>();
  for (const match of matches) {
    const value = Number.parseInt(match[1]!, 10);
    if (Number.isInteger(value)) {
      values.add(value);
    }
  }
  return values;
}

function isSeasonMentioned(titles: string[], season: number | null): boolean {
  if (!season) {
    return false;
  }
  const hints = titles.flatMap((title) => [...extractCandidateSeasonHints(title)]);
  return hints.includes(season);
}

function pickBestSearchResult(
  title: string,
  episode: number | null,
  season: number | null,
  media: AnilistMedia[],
): { id: number; title: string } | null {
  const cleanedTitle = removeSeasonHint(title);
  const targets = [title, cleanedTitle]
    .map(normalizeTitle)
    .map((value) => value.trim())
    .filter((value, index, all) => value.length > 0 && all.indexOf(value) === index);

  const filtered =
    episode === null
      ? media
      : media.filter((item) => {
          const total = item.episodes;
          return total === null || total >= episode;
        });
  const candidates = filtered.length > 0 ? filtered : media;
  if (candidates.length === 0) {
    return null;
  }

  const scored = candidates.map((item) => {
    const candidateTitles = [item.title?.romaji, item.title?.english, item.title?.native]
      .filter((value): value is string => typeof value === 'string')
      .map((value) => normalizeTitle(value));

    let score = 0;

    for (const target of targets) {
      if (candidateTitles.includes(target)) {
        score += 120;
        continue;
      }
      if (candidateTitles.some((itemTitle) => itemTitle.includes(target))) {
        score += 30;
      }
      if (candidateTitles.some((itemTitle) => target.includes(itemTitle))) {
        score += 10;
      }
    }

    if (episode !== null && item.episodes === episode) {
      score += 20;
    }

    if (season !== null && isSeasonMentioned(candidateTitles, season)) {
      score += 15;
    }

    return { item, score };
  });

  scored.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return b.item.id - a.item.id;
  });

  const selected = scored[0]!;
  const selectedTitle =
    selected.item.title?.english ??
    selected.item.title?.romaji ??
    selected.item.title?.native ??
    title;
  return { id: selected.item.id, title: selectedTitle };
}

function buildSearchCandidates(parsed: CoverArtCandidate): string[] {
  const candidateTitles = [
    ...(parsed.source === 'guessit' && parsed.season !== null && parsed.season > 1
      ? [`${parsed.title} Season ${parsed.season}`]
      : []),
    parsed.title,
  ];
  return candidateTitles
    .map((title) => title.trim())
    .filter((title, index, all) => title.length > 0 && all.indexOf(title) === index);
}

async function searchAnilist(
  rateLimiter: AnilistRateLimiter,
  title: string,
): Promise<{ media: AnilistMedia[]; rateLimited: boolean }> {
  await rateLimiter.acquire();

  const res = await fetch(ANILIST_GRAPHQL_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify({ query: SEARCH_QUERY, variables: { search: title } }),
  });

  rateLimiter.recordResponse(res.headers);

  if (res.status === 429) {
    return { media: [], rateLimited: true };
  }

  if (!res.ok) {
    throw new Error(`Anilist search failed: ${res.status} ${res.statusText}`);
  }

  const json = (await res.json()) as AnilistSearchResponse;
  const mediaList = json.data?.Page?.media;
  if (!mediaList || mediaList.length === 0) {
    return { media: [], rateLimited: false };
  }

  return { media: mediaList, rateLimited: false };
}

async function downloadImage(url: string): Promise<Buffer | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const arrayBuf = await res.arrayBuffer();
    return Buffer.from(arrayBuf);
  } catch {
    return null;
  }
}

export function createCoverArtFetcher(
  rateLimiter: AnilistRateLimiter,
  logger: Logger,
  options: CoverArtFetcherOptions = {},
): CoverArtFetcher {
  const reuseAnimeCoverArt = (db: DatabaseSync, videoId: number): boolean => {
    const row = db
      .prepare('SELECT anime_id AS animeId FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { animeId: number | null } | undefined;
    if (!row?.animeId) {
      return false;
    }

    const shared = getAnimeCoverArt(db, row.animeId);
    if (!shared?.coverBlob) {
      return false;
    }

    upsertCoverArt(db, videoId, {
      anilistId: shared.anilistId,
      coverUrl: shared.coverUrl,
      coverBlob: shared.coverBlob,
      titleRomaji: shared.titleRomaji,
      titleEnglish: shared.titleEnglish,
      episodesTotal: shared.episodesTotal,
    });
    return true;
  };

  const resolveCanonicalTitle = (
    db: DatabaseSync,
    videoId: number,
    fallbackTitle: string,
  ): string => {
    const row = db
      .prepare(
        `
          SELECT canonical_title AS canonicalTitle
          FROM imm_videos
          WHERE video_id = ?
          LIMIT 1
        `,
      )
      .get(videoId) as { canonicalTitle: string | null } | undefined;
    return row?.canonicalTitle?.trim() || fallbackTitle;
  };

  const resolveMediaInfo = async (
    db: DatabaseSync,
    videoId: number,
    canonicalTitle: string,
  ): Promise<CoverArtCandidate | null> => {
    const effectiveTitle = resolveCanonicalTitle(db, videoId, canonicalTitle);
    const parsed = await guessAnilistMediaInfo(null, effectiveTitle, {
      runGuessit: options.runGuessit ?? runGuessit,
    });
    if (!parsed) {
      return null;
    }
    return {
      title: parsed.title,
      season: parsed.season,
      episode: parsed.episode,
      source: parsed.source,
    };
  };

  return {
    async fetchIfMissing(db, videoId, canonicalTitle): Promise<boolean> {
      const existing = getCoverArt(db, videoId);
      if (existing?.coverBlob) {
        return true;
      }

      if (existing?.coverUrl) {
        const coverBlob = await downloadImage(existing.coverUrl);
        if (coverBlob) {
          upsertCoverArt(db, videoId, {
            anilistId: existing.anilistId,
            coverUrl: existing.coverUrl,
            coverBlob,
            titleRomaji: existing.titleRomaji,
            titleEnglish: existing.titleEnglish,
            episodesTotal: existing.episodesTotal,
          });
          return true;
        }
      }

      if (reuseAnimeCoverArt(db, videoId)) {
        return true;
      }

      if (
        existing &&
        existing.coverUrl === null &&
        existing.anilistId === null &&
        Date.now() - existing.fetchedAtMs < NO_MATCH_RETRY_MS
      ) {
        return false;
      }

      const effectiveTitle = resolveCanonicalTitle(db, videoId, canonicalTitle);
      const cleaned = stripFilenameTags(effectiveTitle);
      if (!cleaned) {
        logger.warn('cover-art: empty title after stripping tags for videoId=%d', videoId);
        upsertCoverArt(db, videoId, {
          anilistId: null,
          coverUrl: null,
          coverBlob: null,
          titleRomaji: null,
          titleEnglish: null,
          episodesTotal: null,
        });
        return false;
      }

      const parsedInfo = await resolveMediaInfo(db, videoId, canonicalTitle);
      const searchBase = parsedInfo?.title ?? cleaned;
      const searchCandidates = parsedInfo ? buildSearchCandidates(parsedInfo) : [cleaned];

      const effectiveCandidates = searchCandidates.includes(cleaned)
        ? searchCandidates
        : [...searchCandidates, cleaned];

      let selected: AnilistMedia | null = null;
      let rateLimited = false;

      for (const candidate of effectiveCandidates) {
        logger.info('cover-art: searching Anilist for "%s" (videoId=%d)', candidate, videoId);

        try {
          const result = await searchAnilist(rateLimiter, candidate);
          rateLimited = result.rateLimited;
          if (result.media.length === 0) {
            continue;
          }

          const picked = pickBestSearchResult(
            searchBase,
            parsedInfo?.episode ?? null,
            parsedInfo?.season ?? null,
            result.media,
          );
          if (picked) {
            const match = result.media.find((media) => media.id === picked.id);
            if (match) {
              selected = match;
              break;
            }
          }
        } catch (err) {
          logger.error('cover-art: Anilist search error for "%s": %s', candidate, err);
          return false;
        }
      }

      if (rateLimited) {
        logger.warn('cover-art: rate-limited by Anilist, skipping videoId=%d', videoId);
        return false;
      }

      if (!selected) {
        logger.info('cover-art: no Anilist results for "%s", caching no-match', searchBase);
        upsertCoverArt(db, videoId, {
          anilistId: null,
          coverUrl: null,
          coverBlob: null,
          titleRomaji: null,
          titleEnglish: null,
          episodesTotal: null,
        });
        return false;
      }

      const coverUrl = selected.coverImage?.large ?? selected.coverImage?.medium ?? null;
      let coverBlob: Buffer | null = null;
      if (coverUrl) {
        coverBlob = await downloadImage(coverUrl);
      }

      upsertCoverArt(db, videoId, {
        anilistId: selected.id,
        coverUrl,
        coverBlob,
        titleRomaji: selected.title?.romaji ?? null,
        titleEnglish: selected.title?.english ?? null,
        episodesTotal: selected.episodes ?? null,
      });

      updateAnimeAnilistInfo(db, videoId, {
        anilistId: selected.id,
        titleRomaji: selected.title?.romaji ?? null,
        titleEnglish: selected.title?.english ?? null,
        titleNative: selected.title?.native ?? null,
        episodesTotal: selected.episodes ?? null,
      });

      logger.info(
        'cover-art: cached art for videoId=%d anilistId=%d title="%s"',
        videoId,
        selected.id,
        selected.title?.romaji ?? searchBase,
      );

      return true;
    },
  };
}
