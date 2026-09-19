import * as childProcess from 'node:child_process';
import type { TmdbMediaType } from '../../../shared/media-kind';
import type { TmdbConfig } from '../../../types/integrations';
import type { StatsTmdbSearchResult } from '../../../types/stats-http-contract';

export const TMDB_API_BASE_URL = 'https://api.themoviedb.org/3';
const TMDB_POSTER_BASE_URL = 'https://image.tmdb.org/t/p/w500';
const REQUEST_TIMEOUT_MS = 8_000;
const API_KEY_COMMAND_RETRY_MS = 30_000;
const ANIMATION_GENRE_ID = 16;

export type TmdbSearchResult = StatsTmdbSearchResult;

export interface TmdbTitleDetails {
  tmdbId: number;
  tmdbType: TmdbMediaType;
  titleEnglish: string | null;
  titleNative: string | null;
  /** English synopsis, falling back to the Japanese one. */
  description: string | null;
  posterUrl: string | null;
  episodesTotal: number | null;
  year: number | null;
  originalLanguage: string;
  isAnimation: boolean;
  /** Every name TMDB knows for the title, used for exact-title matching. */
  allTitles: string[];
}

export interface TmdbClient {
  search(query: string): Promise<TmdbSearchResult[]>;
  getDetails(tmdbType: TmdbMediaType, tmdbId: number): Promise<TmdbTitleDetails | null>;
}

export class TmdbApiKeyMissingError extends Error {
  constructor() {
    super('TMDB API key not configured. Set tmdb.apiKey or tmdb.apiKeyCommand.');
    this.name = 'TmdbApiKeyMissingError';
  }
}

export class TmdbRequestError extends Error {
  constructor(
    message: string,
    readonly status: number,
  ) {
    super(message);
    this.name = 'TmdbRequestError';
  }
}

function execCommand(command: string): Promise<string> {
  return new Promise((resolve, reject) => {
    childProcess.exec(command, { timeout: 10_000 }, (err, stdout) => {
      if (err) {
        reject(err);
        return;
      }
      resolve(stdout);
    });
  });
}

/**
 * Resolves the key in priority order: the user's literal `apiKey`, then the
 * output of `apiKeyCommand`, then the key bundled into release builds.
 */
export async function resolveTmdbApiKey(
  config: TmdbConfig | undefined,
  bundledKey: string | null = null,
): Promise<string | null> {
  const literal = config?.apiKey?.trim();
  if (literal) return literal;
  const command = config?.apiKeyCommand?.trim();
  if (command) {
    try {
      const key = (await execCommand(command)).trim();
      if (key.length > 0) return key;
    } catch {
      /* fall through to the bundled key */
    }
  }
  return bundledKey;
}

/** Cache successful command output until either credential setting changes. */
export function createTmdbApiKeyResolver(
  getConfig: () => TmdbConfig | undefined,
  getBundledKey: () => string | null = () => null,
): () => Promise<string | null> {
  let state:
    | {
        apiKey: string | undefined;
        apiKeyCommand: string | undefined;
        pending: Promise<string | null> | null;
        retryAfterMs: number;
      }
    | undefined;

  return async () => {
    const config = getConfig();
    if (
      !state ||
      state.apiKey !== config?.apiKey ||
      state.apiKeyCommand !== config?.apiKeyCommand
    ) {
      state = {
        apiKey: config?.apiKey,
        apiKeyCommand: config?.apiKeyCommand,
        pending: null,
        retryAfterMs: 0,
      };
    }
    const current = state;
    const literal = current.apiKey?.trim();
    if (literal) return literal;
    if (!current.apiKeyCommand?.trim()) return getBundledKey();
    if (Date.now() < current.retryAfterMs) return getBundledKey();
    current.pending ??= resolveTmdbApiKey(current).then((key) => {
      if (!key) {
        current.retryAfterMs = Date.now() + API_KEY_COMMAND_RETRY_MS;
        current.pending = null;
      }
      return key;
    });
    const key = await current.pending;
    return key ?? getBundledKey();
  };
}

interface RawSearchItem {
  media_type?: string;
  id?: number;
  name?: string;
  original_name?: string;
  title?: string;
  original_title?: string;
  original_language?: string;
  overview?: string;
  poster_path?: string | null;
  first_air_date?: string;
  release_date?: string;
  genre_ids?: number[];
}

interface RawTranslation {
  iso_639_1?: string;
  data?: { name?: string; title?: string; overview?: string };
}

interface RawDetails extends RawSearchItem {
  number_of_episodes?: number;
  genres?: Array<{ id?: number }>;
  alternative_titles?: { results?: Array<{ title?: string }>; titles?: Array<{ title?: string }> };
  translations?: { translations?: RawTranslation[] };
}

function nonEmpty(value: unknown): string | null {
  return typeof value === 'string' && value.trim() ? value.trim() : null;
}

function yearOf(date: string | undefined): number | null {
  const year = Number.parseInt(date?.slice(0, 4) ?? '', 10);
  return Number.isFinite(year) && year > 0 ? year : null;
}

function posterUrlOf(path: string | null | undefined): string | null {
  return path ? `${TMDB_POSTER_BASE_URL}${path}` : null;
}

function mediaTypeOf(value: unknown): TmdbMediaType | null {
  return value === 'tv' || value === 'movie' ? value : null;
}

function normalizeSearchItem(item: RawSearchItem): TmdbSearchResult | null {
  const tmdbType = mediaTypeOf(item.media_type);
  if (!tmdbType || typeof item.id !== 'number') return null;
  const title = nonEmpty(item.name) ?? nonEmpty(item.title);
  const originalTitle = nonEmpty(item.original_name) ?? nonEmpty(item.original_title) ?? title;
  if (!title || !originalTitle) return null;
  return {
    tmdbId: item.id,
    tmdbType,
    title,
    originalTitle,
    originalLanguage: item.original_language ?? '',
    overview: nonEmpty(item.overview),
    posterUrl: posterUrlOf(item.poster_path),
    year: yearOf(item.first_air_date ?? item.release_date),
    isAnimation: (item.genre_ids ?? []).includes(ANIMATION_GENRE_ID),
  };
}

function normalizeDetails(tmdbType: TmdbMediaType, raw: RawDetails): TmdbTitleDetails | null {
  if (typeof raw.id !== 'number') return null;
  const localizedTitle = nonEmpty(raw.name) ?? nonEmpty(raw.title);
  const originalTitle = nonEmpty(raw.original_name) ?? nonEmpty(raw.original_title);
  const originalLanguage = raw.original_language ?? '';
  const translations = raw.translations?.translations ?? [];
  const translationFor = (language: string) =>
    translations.find((entry) => entry.iso_639_1 === language)?.data;
  const english = translationFor('en');
  const japanese = translationFor('ja');
  const englishTitle =
    nonEmpty(english?.name) ??
    nonEmpty(english?.title) ??
    (localizedTitle && localizedTitle !== originalTitle ? localizedTitle : null);
  const nativeTitle =
    originalLanguage === 'ja'
      ? originalTitle
      : (nonEmpty(japanese?.name) ?? nonEmpty(japanese?.title));
  const alternativeTitles = [
    ...(raw.alternative_titles?.results ?? []),
    ...(raw.alternative_titles?.titles ?? []),
  ].map((entry) => nonEmpty(entry.title));
  const translatedTitles = translations.map(
    (entry) => nonEmpty(entry.data?.name) ?? nonEmpty(entry.data?.title),
  );
  const allTitles = [
    ...new Set(
      [
        localizedTitle,
        originalTitle,
        englishTitle,
        nativeTitle,
        ...translatedTitles,
        ...alternativeTitles,
      ].filter((title): title is string => Boolean(title)),
    ),
  ];
  return {
    tmdbId: raw.id,
    tmdbType,
    titleEnglish: englishTitle,
    titleNative: nativeTitle,
    description:
      nonEmpty(raw.overview) ?? nonEmpty(english?.overview) ?? nonEmpty(japanese?.overview),
    posterUrl: posterUrlOf(raw.poster_path),
    episodesTotal:
      tmdbType === 'movie'
        ? 1
        : typeof raw.number_of_episodes === 'number' && raw.number_of_episodes > 0
          ? raw.number_of_episodes
          : null,
    year: yearOf(raw.first_air_date ?? raw.release_date),
    originalLanguage,
    isAnimation: (raw.genres ?? []).some((genre) => genre.id === ANIMATION_GENRE_ID),
    allTitles,
  };
}

// TMDB issues two kinds of credential: a short v3 key that travels as a query
// parameter and a long v4 read token (a JWT) that goes in the Authorization
// header. Users paste whichever the settings page showed them.
function isV4Token(apiKey: string): boolean {
  return apiKey.startsWith('eyJ');
}

export function createTmdbClient(deps: {
  resolveApiKey: () => Promise<string | null>;
  fetch?: typeof fetch;
  baseUrl?: string;
}): TmdbClient {
  const fetchImpl = deps.fetch ?? fetch;
  const baseUrl = (deps.baseUrl ?? TMDB_API_BASE_URL).replace(/\/+$/, '');

  async function request<T>(path: string, params: Record<string, string>): Promise<T | null> {
    const apiKey = await deps.resolveApiKey();
    if (!apiKey) throw new TmdbApiKeyMissingError();
    const url = new URL(`${baseUrl}${path}`);
    for (const [key, value] of Object.entries(params)) url.searchParams.set(key, value);
    const headers: Record<string, string> = { Accept: 'application/json' };
    if (isV4Token(apiKey)) {
      headers.Authorization = `Bearer ${apiKey}`;
    } else {
      url.searchParams.set('api_key', apiKey);
    }
    const res = await fetchImpl(url, { headers, signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
    if (res.status === 404) return null;
    if (!res.ok) {
      throw new TmdbRequestError(
        `TMDB request failed: ${res.status} ${res.statusText}`,
        res.status,
      );
    }
    return (await res.json()) as T;
  }

  return {
    async search(query) {
      const trimmed = query.trim();
      if (!trimmed) return [];
      const payload = await request<{ results?: RawSearchItem[] }>('/search/multi', {
        query: trimmed,
        include_adult: 'false',
        language: 'en-US',
        page: '1',
      });
      return (payload?.results ?? [])
        .map(normalizeSearchItem)
        .filter((item): item is TmdbSearchResult => item !== null);
    },
    async getDetails(tmdbType, tmdbId) {
      const raw = await request<RawDetails>(`/${tmdbType}/${tmdbId}`, {
        language: 'en-US',
        append_to_response: 'alternative_titles,translations',
      });
      return raw ? normalizeDetails(tmdbType, raw) : null;
    },
  };
}
