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
  type AnilistMediaGuess,
  type GuessAnilistMediaInfoDeps,
} from './anilist-updater';
import {
  resolveAnilistSeasonMedia,
  type AnilistQueryExecutor,
  type AnilistSeasonResolution,
} from './season-resolver';

const ANILIST_GRAPHQL_URL = 'https://graphql.anilist.co';
const NO_MATCH_RETRY_MS = 5 * 60 * 1000;

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
  source: AnilistMediaGuess['source'];
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

class AnilistRateLimitedError extends Error {
  constructor() {
    super('Anilist rate limit reached');
    this.name = 'AnilistRateLimitedError';
  }
}

async function executeAnilistQuery<T>(
  rateLimiter: AnilistRateLimiter,
  query: string,
  variables: Record<string, unknown>,
): Promise<T> {
  await rateLimiter.acquire();

  const res = await fetch(ANILIST_GRAPHQL_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify({ query, variables }),
  });

  rateLimiter.recordResponse(res.headers);

  if (res.status === 429) {
    throw new AnilistRateLimitedError();
  }

  if (!res.ok) {
    throw new Error(`Anilist search failed: ${res.status} ${res.statusText}`);
  }

  const json = (await res.json()) as { data?: T; errors?: Array<{ message?: string }> };
  const firstError = json.errors?.find((entry) => Boolean(entry?.message));
  if (firstError?.message) {
    throw new Error(firstError.message);
  }
  if (!json.data) {
    throw new Error('Anilist response missing data');
  }
  return json.data;
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
      const searchTitles = searchBase === cleaned ? [searchBase] : ([searchBase, cleaned] as const);

      const execute: AnilistQueryExecutor = (query, variables) =>
        executeAnilistQuery(rateLimiter, query, variables);

      let resolution: AnilistSeasonResolution | null = null;
      try {
        for (const searchTitle of searchTitles) {
          logger.info('cover-art: searching Anilist for "%s" (videoId=%d)', searchTitle, videoId);
          resolution = await resolveAnilistSeasonMedia(
            {
              title: searchTitle,
              season: parsedInfo?.season ?? null,
              episode: parsedInfo?.episode ?? null,
            },
            { execute, logInfo: (message) => logger.info('%s', message) },
          );
          if (resolution) break;
        }
      } catch (err) {
        if (err instanceof AnilistRateLimitedError) {
          logger.warn('cover-art: rate-limited by Anilist, skipping videoId=%d', videoId);
          return false;
        }
        logger.error('cover-art: Anilist search error for "%s": %s', searchBase, err);
        return false;
      }

      if (resolution && !resolution.seasonResolved) {
        // Only the season 1 entry was found. Storing its artwork would leave a cover with
        // no AniList id, which the `existing.coverBlob` early return above serves forever,
        // so the season could never re-resolve once AniList publishes the relation.
        // Caching a plain no-match instead reuses the NO_MATCH_RETRY_MS retry window.
        logger.warn(
          'cover-art: could not find season %d of "%s" (only matched "%s"), caching no-match',
          resolution.requestedSeason,
          searchBase,
          resolution.title,
        );
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

      const selected = resolution?.media ?? null;
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
        exactTitleMatch: resolution?.exactTitleMatch ?? false,
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
