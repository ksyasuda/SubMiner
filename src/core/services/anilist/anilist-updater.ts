import * as childProcess from 'child_process';
import * as path from 'path';

import { parseMediaInfo } from '../../../jimaku/utils';
import type { AnilistRateLimiter } from './rate-limiter';
import { resolveAnilistSeasonMedia } from './season-resolver';

const ANILIST_GRAPHQL_URL = 'https://graphql.anilist.co';

export interface AnilistMediaGuess {
  title: string;
  alternativeTitle?: string;
  year?: number;
  season: number | null;
  episode: number | null;
  /** `stream` means the player was handed these fields, not that we parsed them. */
  source: 'guessit' | 'fallback' | 'stream';
}

export interface AnilistPostWatchUpdateResult {
  status: 'updated' | 'skipped' | 'error';
  message: string;
  retryable?: boolean;
}

export interface AnilistPostWatchUpdateOptions {
  rateLimiter?: AnilistRateLimiter;
  season?: number | null;
  /**
   * Pinned AniList media id (from a character dictionary manual override). When set,
   * the search/season resolution is skipped entirely.
   */
  mediaId?: number | null;
  logInfo?: (message: string) => void;
}

interface AnilistGraphQlError {
  message?: string;
}

interface AnilistGraphQlResponse<T> {
  data?: T;
  errors?: AnilistGraphQlError[];
}

interface AnilistMediaEntryData {
  Media?: {
    id: number;
    episodes?: number | null;
    title?: {
      romaji?: string | null;
      english?: string | null;
      native?: string | null;
    } | null;
    mediaListEntry?: {
      progress?: number | null;
      status?: string | null;
    } | null;
  } | null;
}

interface AnilistSaveEntryData {
  SaveMediaListEntry?: {
    progress?: number | null;
    status?: string | null;
  };
}

export function runGuessit(target: string): Promise<string> {
  return new Promise((resolve, reject) => {
    childProcess.execFile(
      'guessit',
      [target, '--json'],
      { timeout: 5000, maxBuffer: 1024 * 1024 },
      (error, stdout) => {
        if (error) {
          reject(error);
          return;
        }
        resolve(stdout);
      },
    );
  });
}

export interface GuessAnilistMediaInfoDeps {
  runGuessit: (target: string) => Promise<string>;
}

function firstString(value: unknown): string | null {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : null;
  }
  if (Array.isArray(value)) {
    for (const item of value) {
      const candidate = firstString(item);
      if (candidate) return candidate;
    }
  }
  return null;
}

function normalizeGuessitTitlePart(value: string): string {
  return value.replace(/[._]+/g, ' ').replace(/-/g, ' ').replace(/\s+/g, ' ').trim();
}

function readGuessitTitle(value: unknown): string | null {
  if (typeof value === 'string') {
    const normalized = normalizeGuessitTitlePart(value);
    return normalized.length > 0 ? normalized : null;
  }
  if (Array.isArray(value)) {
    const parts = value
      .filter((item): item is string => typeof item === 'string')
      .map((item) => normalizeGuessitTitlePart(item))
      .filter((item) => item.length > 0);
    if (parts.length === 0) {
      return null;
    }
    return parts.join(' ').replace(/\s+/g, ' ').trim();
  }
  return null;
}

function firstPositiveInteger(value: unknown): number | null {
  if (typeof value === 'number' && Number.isInteger(value) && value > 0) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number.parseInt(value, 10);
    return Number.isInteger(parsed) && parsed > 0 ? parsed : null;
  }
  if (Array.isArray(value)) {
    for (const item of value) {
      const candidate = firstPositiveInteger(item);
      if (candidate !== null) return candidate;
    }
  }
  return null;
}

function firstYear(value: unknown): number | undefined {
  const candidate = firstPositiveInteger(value);
  if (candidate === null) return undefined;
  return candidate >= 1900 && candidate <= 2200 ? candidate : undefined;
}

function buildGuessitTitle(title: string, alternativeTitle: string | null): string {
  if (!alternativeTitle) return title;
  if (title.length <= 3) {
    return `${title} ${alternativeTitle}`.replace(/\s+/g, ' ').trim();
  }
  return title;
}

async function anilistGraphQl<T>(
  accessToken: string,
  query: string,
  variables: Record<string, unknown>,
  options: AnilistPostWatchUpdateOptions = {},
): Promise<AnilistGraphQlResponse<T>> {
  try {
    await options.rateLimiter?.acquire();
    const response = await fetch(ANILIST_GRAPHQL_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${accessToken}`,
      },
      body: JSON.stringify({ query, variables }),
    });

    options.rateLimiter?.recordResponse(response.headers);
    const payload = (await response.json()) as AnilistGraphQlResponse<T>;
    return payload;
  } catch (error) {
    return {
      errors: [
        {
          message: error instanceof Error ? error.message : String(error),
        },
      ],
    };
  }
}

function firstErrorMessage<T>(response: AnilistGraphQlResponse<T>): string | null {
  const firstError = response.errors?.find((item) => Boolean(item?.message));
  return firstError?.message ?? null;
}

/** The season resolver signals failure by throwing; anilistGraphQl reports it in-band. */
function createAnilistSeasonQueryExecutor(
  accessToken: string,
  options: AnilistPostWatchUpdateOptions,
) {
  return async <T>(query: string, variables: Record<string, unknown>): Promise<T> => {
    const response = await anilistGraphQl<T>(accessToken, query, variables, options);
    const error = firstErrorMessage(response);
    if (error) {
      throw new Error(error);
    }
    if (!response.data) {
      throw new Error('AniList response missing data');
    }
    return response.data;
  };
}

function isUpdateableListStatus(status: string | null | undefined): boolean {
  return status === 'CURRENT' || status === 'PLANNING';
}

function formatListStatus(status: string | null | undefined): string {
  if (!status) return 'not in your AniList Planning or Watching list';
  return `marked ${status.toLowerCase().replace(/_/g, ' ')} on AniList`;
}

function isKnownFinalEpisode(totalEpisodes: number | null, episode: number): boolean {
  return (
    typeof totalEpisodes === 'number' &&
    Number.isInteger(totalEpisodes) &&
    totalEpisodes > 0 &&
    episode === totalEpisodes
  );
}

export async function guessAnilistMediaInfo(
  mediaPath: string | null,
  mediaTitle: string | null,
  deps: GuessAnilistMediaInfoDeps = { runGuessit },
): Promise<AnilistMediaGuess | null> {
  const target = mediaPath ?? mediaTitle;
  const guessitTarget = mediaPath ? path.basename(mediaPath) : mediaTitle;

  if (guessitTarget && guessitTarget.trim().length > 0) {
    try {
      const stdout = await deps.runGuessit(guessitTarget);
      const parsed = JSON.parse(stdout) as Record<string, unknown>;
      const title = readGuessitTitle(parsed.title);
      const alternativeTitle = readGuessitTitle(parsed.alternative_title);
      const episode = firstPositiveInteger(parsed.episode);
      const season = firstPositiveInteger(parsed.season);
      const year = firstYear(parsed.year);
      if (title) {
        const fallback = parseMediaInfo(target);
        const canUseFallbackDetails = fallback.confidence !== 'low';
        return {
          title: buildGuessitTitle(title, alternativeTitle),
          ...(alternativeTitle ? { alternativeTitle } : {}),
          ...(year ? { year } : {}),
          season: season ?? fallback.season,
          episode: episode ?? (canUseFallbackDetails ? fallback.episode : null),
          source: 'guessit',
        };
      }
    } catch {
      // Ignore guessit failures and fall back to internal parser.
    }
  }

  const fallbackTarget = mediaPath ?? mediaTitle;
  const parsed = parseMediaInfo(fallbackTarget);
  if (!parsed.title.trim()) {
    return null;
  }
  return {
    title: parsed.title.trim(),
    season: parsed.season,
    episode: parsed.episode,
    source: 'fallback',
  };
}

export async function updateAnilistPostWatchProgress(
  accessToken: string,
  title: string,
  episode: number,
  options: AnilistPostWatchUpdateOptions = {},
): Promise<AnilistPostWatchUpdateResult> {
  const pinnedMediaId =
    typeof options.mediaId === 'number' && Number.isInteger(options.mediaId) && options.mediaId > 0
      ? options.mediaId
      : null;

  let mediaId = pinnedMediaId;
  let resolvedTitle: string | null = null;
  let resolvedEpisodes: number | null = null;

  if (mediaId === null) {
    let resolution: Awaited<ReturnType<typeof resolveAnilistSeasonMedia>>;
    try {
      resolution = await resolveAnilistSeasonMedia(
        { title, season: options.season, episode },
        {
          execute: createAnilistSeasonQueryExecutor(accessToken, options),
          logInfo: options.logInfo,
        },
      );
    } catch (error) {
      return {
        status: 'error',
        message: `AniList search failed: ${error instanceof Error ? error.message : String(error)}`,
      };
    }

    if (!resolution) {
      // A well-formed search that matched nothing is deterministic for this title, so
      // requeueing it just burns rate limit. Transient failures throw and are caught above.
      return {
        status: 'error',
        retryable: false,
        message: 'AniList search returned no matches.',
      };
    }
    if (!resolution.seasonResolved) {
      // Updating the season 1 entry here is worse than not updating at all.
      return {
        status: 'error',
        retryable: false,
        message: `AniList update skipped: could not find season ${resolution.requestedSeason} of "${title}" (only matched "${resolution.title}"). Pick the right entry with the character dictionary AniList override.`,
      };
    }

    mediaId = resolution.id;
    resolvedTitle = resolution.title;
    resolvedEpisodes = resolution.episodes;
  }

  const entryResponse = await anilistGraphQl<AnilistMediaEntryData>(
    accessToken,
    `
      query ($mediaId: Int!) {
        Media(id: $mediaId, type: ANIME) {
          id
          episodes
          title {
            romaji
            english
            native
          }
          mediaListEntry {
            progress
            status
          }
        }
      }
    `,
    { mediaId },
    options,
  );
  const entryError = firstErrorMessage(entryResponse);
  if (entryError) {
    return {
      status: 'error',
      message: `AniList entry lookup failed: ${entryError}`,
    };
  }

  const entryMedia = entryResponse.data?.Media ?? null;
  const pickedTitle =
    resolvedTitle ||
    entryMedia?.title?.english?.trim() ||
    entryMedia?.title?.romaji?.trim() ||
    entryMedia?.title?.native?.trim() ||
    title;
  const pickedEpisodes =
    resolvedEpisodes ??
    (typeof entryMedia?.episodes === 'number' && entryMedia.episodes > 0
      ? entryMedia.episodes
      : null);

  const entry = entryMedia?.mediaListEntry ?? null;
  if (!entry || !isUpdateableListStatus(entry.status)) {
    return {
      status: 'error',
      retryable: false,
      message: `AniList update not possible: "${pickedTitle}" is ${formatListStatus(entry?.status)}. Add it to Planning or Watching, then mark watched again.`,
    };
  }

  const currentProgress = entry.progress ?? 0;
  const shouldMarkCompleted = isKnownFinalEpisode(pickedEpisodes, episode);
  if (typeof currentProgress === 'number' && currentProgress >= episode && !shouldMarkCompleted) {
    return {
      status: 'skipped',
      message: `AniList already at episode ${currentProgress} (${pickedTitle}).`,
    };
  }

  const saveResponse = await anilistGraphQl<AnilistSaveEntryData>(
    accessToken,
    `
      mutation ($mediaId: Int!, $progress: Int!, $status: MediaListStatus!) {
        SaveMediaListEntry(mediaId: $mediaId, progress: $progress, status: $status) {
          progress
          status
        }
      }
    `,
    {
      mediaId,
      progress: episode,
      status: shouldMarkCompleted ? 'COMPLETED' : 'CURRENT',
    },
    options,
  );
  const saveError = firstErrorMessage(saveResponse);
  if (saveError) {
    return { status: 'error', message: `AniList update failed: ${saveError}` };
  }

  return {
    status: 'updated',
    message: shouldMarkCompleted
      ? `AniList updated "${pickedTitle}" to episode ${episode} and marked it completed.`
      : `AniList updated "${pickedTitle}" to episode ${episode}.`,
  };
}
