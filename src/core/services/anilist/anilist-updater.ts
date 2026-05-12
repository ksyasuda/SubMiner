import * as childProcess from 'child_process';
import * as path from 'path';

import { parseMediaInfo } from '../../../jimaku/utils';

const ANILIST_GRAPHQL_URL = 'https://graphql.anilist.co';

export interface AnilistMediaGuess {
  title: string;
  alternativeTitle?: string;
  year?: number;
  season: number | null;
  episode: number | null;
  source: 'guessit' | 'fallback';
}

export interface AnilistPostWatchUpdateResult {
  status: 'updated' | 'skipped' | 'error';
  message: string;
}

interface AnilistGraphQlError {
  message?: string;
}

interface AnilistGraphQlResponse<T> {
  data?: T;
  errors?: AnilistGraphQlError[];
}

interface AnilistSearchData {
  Page?: {
    media?: Array<{
      id: number;
      episodes: number | null;
      title?: {
        romaji?: string | null;
        english?: string | null;
        native?: string | null;
      };
    }>;
  };
}

interface AnilistMediaEntryData {
  Media?: {
    id: number;
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

function normalizeTitle(text: string): string {
  return text.trim().toLowerCase().replace(/\s+/g, ' ');
}

async function anilistGraphQl<T>(
  accessToken: string,
  query: string,
  variables: Record<string, unknown>,
): Promise<AnilistGraphQlResponse<T>> {
  try {
    const response = await fetch(ANILIST_GRAPHQL_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${accessToken}`,
      },
      body: JSON.stringify({ query, variables }),
    });

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

function pickBestSearchResult(
  title: string,
  episode: number,
  media: Array<{
    id: number;
    episodes: number | null;
    title?: {
      romaji?: string | null;
      english?: string | null;
      native?: string | null;
    };
  }>,
): { id: number; title: string } | null {
  const filtered = media.filter((item) => {
    const totalEpisodes = item.episodes;
    return totalEpisodes === null || totalEpisodes >= episode;
  });
  const candidates = filtered.length > 0 ? filtered : media;
  if (candidates.length === 0) return null;

  const normalizedTarget = normalizeTitle(title);
  const exact = candidates.find((item) => {
    const titles = [item.title?.romaji, item.title?.english, item.title?.native]
      .filter((value): value is string => typeof value === 'string')
      .map((value) => normalizeTitle(value));
    return titles.includes(normalizedTarget);
  });

  const selected = exact ?? candidates[0]!;
  const selectedTitle =
    selected.title?.english || selected.title?.romaji || selected.title?.native || title;
  return { id: selected.id, title: selectedTitle };
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
          season: season ?? (canUseFallbackDetails ? fallback.season : null),
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
): Promise<AnilistPostWatchUpdateResult> {
  const searchResponse = await anilistGraphQl<AnilistSearchData>(
    accessToken,
    `
      query ($search: String!) {
        Page(perPage: 5) {
          media(search: $search, type: ANIME) {
            id
            episodes
            title {
              romaji
              english
              native
            }
          }
        }
      }
    `,
    { search: title },
  );
  const searchError = firstErrorMessage(searchResponse);
  if (searchError) {
    return {
      status: 'error',
      message: `AniList search failed: ${searchError}`,
    };
  }

  const media = searchResponse.data?.Page?.media ?? [];
  const picked = pickBestSearchResult(title, episode, media);
  if (!picked) {
    return { status: 'error', message: 'AniList search returned no matches.' };
  }

  const entryResponse = await anilistGraphQl<AnilistMediaEntryData>(
    accessToken,
    `
      query ($mediaId: Int!) {
        Media(id: $mediaId, type: ANIME) {
          id
          mediaListEntry {
            progress
            status
          }
        }
      }
    `,
    { mediaId: picked.id },
  );
  const entryError = firstErrorMessage(entryResponse);
  if (entryError) {
    return {
      status: 'error',
      message: `AniList entry lookup failed: ${entryError}`,
    };
  }

  const currentProgress = entryResponse.data?.Media?.mediaListEntry?.progress ?? 0;
  if (typeof currentProgress === 'number' && currentProgress >= episode) {
    return {
      status: 'skipped',
      message: `AniList already at episode ${currentProgress} (${picked.title}).`,
    };
  }

  const saveResponse = await anilistGraphQl<AnilistSaveEntryData>(
    accessToken,
    `
      mutation ($mediaId: Int!, $progress: Int!) {
        SaveMediaListEntry(mediaId: $mediaId, progress: $progress, status: CURRENT) {
          progress
          status
        }
      }
    `,
    { mediaId: picked.id, progress: episode },
  );
  const saveError = firstErrorMessage(saveResponse);
  if (saveError) {
    return { status: 'error', message: `AniList update failed: ${saveError}` };
  }

  return {
    status: 'updated',
    message: `AniList updated "${picked.title}" to episode ${episode}.`,
  };
}
