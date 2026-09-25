import { normalizeTitleIdentity } from '../../utils/title-normalization';
import type { TmdbMediaType } from '../../../shared/media-kind';
import { TmdbApiKeyMissingError, type TmdbClient, type TmdbTitleDetails } from './tmdb-client';

const MAX_DETAIL_LOOKUPS = 3;

/**
 * Resolves live-action titles for the automatic cover-art path. Anime is
 * AniList's job, so only non-animated Japanese-language results qualify, and
 * a candidate must match the parsed title exactly under one of the names TMDB
 * knows for it. Fuzzy search hits are never trusted on their own: a stray
 * filename would otherwise pin the wrong show to a library entry.
 */
export interface LiveActionMetadataResolver {
  resolveByTitle(title: string): Promise<TmdbTitleDetails | null>;
  resolveById(tmdbType: TmdbMediaType, tmdbId: number): Promise<TmdbTitleDetails | null>;
}

interface Logger {
  info(msg: string, ...args: unknown[]): void;
  warn(msg: string, ...args: unknown[]): void;
}

export function titlesMatch(candidate: string, knownTitles: Iterable<string>): boolean {
  const key = normalizeTitleIdentity(candidate);
  if (!key) return false;
  for (const known of knownTitles) {
    if (normalizeTitleIdentity(known) === key) return true;
  }
  return false;
}

export function createLiveActionMetadataResolver(
  client: TmdbClient,
  logger: Logger,
): LiveActionMetadataResolver {
  let warnedMissingKey = false;

  const guard = async <T>(work: () => Promise<T>): Promise<T | null> => {
    try {
      return await work();
    } catch (err) {
      if (err instanceof TmdbApiKeyMissingError) {
        if (!warnedMissingKey) {
          warnedMissingKey = true;
          logger.info('tmdb: no API key configured, skipping live-action metadata lookups');
        }
        return null;
      }
      logger.warn('tmdb: lookup failed: %s', err instanceof Error ? err.message : String(err));
      return null;
    }
  };

  return {
    resolveByTitle(title) {
      return guard(async () => {
        const results = await client.search(title);
        const candidates = results
          .filter((result) => result.originalLanguage === 'ja' && !result.isAnimation)
          .slice(0, MAX_DETAIL_LOOKUPS);
        for (const candidate of candidates) {
          const details = await client.getDetails(candidate.tmdbType, candidate.tmdbId);
          if (details && !details.isAnimation && titlesMatch(title, details.allTitles)) {
            return details;
          }
        }
        return null;
      });
    },
    resolveById(tmdbType, tmdbId) {
      return guard(() => client.getDetails(tmdbType, tmdbId));
    },
  };
}
