import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import https from 'node:https';
import { spawnSync } from 'node:child_process';
import type { Args, JimakuLanguagePreference } from './types.js';
import { DEFAULT_JIMAKU_API_BASE_URL } from './types.js';
import { commandExists } from './util.js';

export interface JimakuEntry {
  id: number;
  name: string;
  english_name?: string | null;
  japanese_name?: string | null;
  flags?: {
    anime?: boolean;
    movie?: boolean;
    adult?: boolean;
    external?: boolean;
    unverified?: boolean;
  };
}

interface JimakuFileEntry {
  name: string;
  url: string;
  size: number;
  last_modified: string;
}

interface JimakuApiError {
  error: string;
  code?: number;
  retryAfter?: number;
}

type JimakuApiResponse<T> = { ok: true; data: T } | { ok: false; error: JimakuApiError };

type JimakuDownloadResult = { ok: true; path: string } | { ok: false; error: JimakuApiError };

interface JimakuConfig {
  apiKey: string;
  apiKeyCommand: string;
  apiBaseUrl: string;
  languagePreference: JimakuLanguagePreference;
  maxEntryResults: number;
}

interface JimakuMediaInfo {
  title: string;
  season: number | null;
  episode: number | null;
  confidence: 'high' | 'medium' | 'low';
  filename: string;
  rawTitle: string;
}

function getRetryAfter(headers: http.IncomingHttpHeaders): number | undefined {
  const value = headers['x-ratelimit-reset-after'];
  if (!value) return undefined;
  const raw = Array.isArray(value) ? value[0] : value;
  const parsed = Number.parseFloat(raw);
  if (!Number.isFinite(parsed)) return undefined;
  return parsed;
}

export function matchEpisodeFromName(name: string): {
  season: number | null;
  episode: number | null;
  index: number | null;
  confidence: 'high' | 'medium' | 'low';
} {
  const seasonEpisode = name.match(/S(\d{1,2})E(\d{1,3})/i);
  if (seasonEpisode && seasonEpisode.index !== undefined) {
    return {
      season: Number.parseInt(seasonEpisode[1], 10),
      episode: Number.parseInt(seasonEpisode[2], 10),
      index: seasonEpisode.index,
      confidence: 'high',
    };
  }

  const alt = name.match(/(\d{1,2})x(\d{1,3})/i);
  if (alt && alt.index !== undefined) {
    return {
      season: Number.parseInt(alt[1], 10),
      episode: Number.parseInt(alt[2], 10),
      index: alt.index,
      confidence: 'high',
    };
  }

  const epOnly = name.match(/(?:^|[\s._-])E(?:P)?(\d{1,3})(?:\b|[\s._-])/i);
  if (epOnly && epOnly.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(epOnly[1], 10),
      index: epOnly.index,
      confidence: 'medium',
    };
  }

  const numeric = name.match(/(?:^|[-–—]\s*)(\d{1,3})\s*[-–—]/);
  if (numeric && numeric.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(numeric[1], 10),
      index: numeric.index,
      confidence: 'medium',
    };
  }

  return { season: null, episode: null, index: null, confidence: 'low' };
}

function detectSeasonFromDir(mediaPath: string): number | null {
  const parent = path.basename(path.dirname(mediaPath));
  const match = parent.match(/(?:Season|S)\s*(\d{1,2})/i);
  if (!match) return null;
  const parsed = Number.parseInt(match[1], 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseGuessitOutput(mediaPath: string, stdout: string): JimakuMediaInfo | null {
  const payload = stdout.trim();
  if (!payload) return null;

  try {
    const parsed = JSON.parse(payload) as {
      title?: string;
      title_original?: string;
      series?: string;
      season?: number | string;
      episode?: number | string;
      episode_list?: Array<number | string>;
    };
    const season =
      typeof parsed.season === 'number'
        ? parsed.season
        : typeof parsed.season === 'string'
          ? Number.parseInt(parsed.season, 10)
          : null;
    const directEpisode =
      typeof parsed.episode === 'number'
        ? parsed.episode
        : typeof parsed.episode === 'string'
          ? Number.parseInt(parsed.episode, 10)
          : null;
    const episodeFromList =
      parsed.episode_list && parsed.episode_list.length > 0
        ? Number.parseInt(String(parsed.episode_list[0]), 10)
        : null;
    const episodeValue =
      directEpisode !== null && Number.isFinite(directEpisode) ? directEpisode : episodeFromList;
    const episode = Number.isFinite(episodeValue as number) ? (episodeValue as number) : null;
    const title = (parsed.title || parsed.title_original || parsed.series || '').trim();
    const hasStructuredData =
      title.length > 0 ||
      Number.isFinite(season as number) ||
      Number.isFinite(episodeValue as number);

    if (!hasStructuredData) return null;

    return {
      title: title || '',
      season: Number.isFinite(season as number) ? season : detectSeasonFromDir(mediaPath),
      episode: episode,
      confidence: 'high',
      filename: path.basename(mediaPath),
      rawTitle: path.basename(mediaPath).replace(/\.[^/.]+$/, ''),
    };
  } catch {
    return null;
  }
}

function parseMediaInfoWithGuessit(mediaPath: string): JimakuMediaInfo | null {
  if (!commandExists('guessit')) return null;

  try {
    const fileName = path.basename(mediaPath);
    const result = spawnSync('guessit', ['--json', fileName], {
      cwd: path.dirname(mediaPath),
      encoding: 'utf8',
      maxBuffer: 2_000_000,
      windowsHide: true,
    });
    if (result.error || result.status !== 0) return null;
    return parseGuessitOutput(mediaPath, result.stdout || '');
  } catch {
    return null;
  }
}

function cleanupTitle(value: string): string {
  return value
    .replace(/^[\s-–—]+/, '')
    .replace(/[\s-–—]+$/, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function formatLangScore(name: string, pref: JimakuLanguagePreference): number {
  if (pref === 'none') return 0;
  const upper = name.toUpperCase();
  const hasJa =
    /(^|[\W_])JA([\W_]|$)/.test(upper) ||
    /(^|[\W_])JPN([\W_]|$)/.test(upper) ||
    upper.includes('.JA.');
  const hasEn =
    /(^|[\W_])EN([\W_]|$)/.test(upper) ||
    /(^|[\W_])ENG([\W_]|$)/.test(upper) ||
    upper.includes('.EN.');
  if (pref === 'ja') {
    if (hasJa) return 2;
    if (hasEn) return 1;
  } else if (pref === 'en') {
    if (hasEn) return 2;
    if (hasJa) return 1;
  }
  return 0;
}

export async function resolveJimakuApiKey(config: JimakuConfig): Promise<string | null> {
  if (config.apiKey && config.apiKey.trim()) {
    return config.apiKey.trim();
  }
  if (config.apiKeyCommand && config.apiKeyCommand.trim()) {
    try {
      const commandResult = spawnSync(config.apiKeyCommand, {
        shell: true,
        encoding: 'utf8',
        timeout: 10000,
      });
      if (commandResult.error) return null;
      const key = (commandResult.stdout || '').trim();
      return key.length > 0 ? key : null;
    } catch {
      return null;
    }
  }
  return null;
}

export function jimakuFetchJson<T>(
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined>,
  options: { baseUrl: string; apiKey: string },
): Promise<JimakuApiResponse<T>> {
  const url = new URL(endpoint, options.baseUrl);
  for (const [key, value] of Object.entries(query)) {
    if (value === null || value === undefined) continue;
    url.searchParams.set(key, String(value));
  }

  return new Promise((resolve) => {
    const requestUrl = new URL(url.toString());
    const transport = requestUrl.protocol === 'https:' ? https : http;
    const req = transport.request(
      requestUrl,
      {
        method: 'GET',
        headers: {
          Authorization: options.apiKey,
          'User-Agent': 'SubMiner',
        },
      },
      (res) => {
        let data = '';
        res.on('data', (chunk) => {
          data += chunk.toString();
        });
        res.on('end', () => {
          const status = res.statusCode || 0;
          if (status >= 200 && status < 300) {
            try {
              const parsed = JSON.parse(data) as T;
              resolve({ ok: true, data: parsed });
            } catch {
              resolve({
                ok: false,
                error: { error: 'Failed to parse Jimaku response JSON.' },
              });
            }
            return;
          }

          let errorMessage = `Jimaku API error (HTTP ${status})`;
          try {
            const parsed = JSON.parse(data) as { error?: string };
            if (parsed && parsed.error) {
              errorMessage = parsed.error;
            }
          } catch {
            // ignore parse errors
          }

          resolve({
            ok: false,
            error: {
              error: errorMessage,
              code: status || undefined,
              retryAfter: status === 429 ? getRetryAfter(res.headers) : undefined,
            },
          });
        });
      },
    );

    req.on('error', (error) => {
      resolve({
        ok: false,
        error: { error: `Jimaku request failed: ${(error as Error).message}` },
      });
    });

    req.end();
  });
}

export function parseMediaInfo(mediaPath: string | null): JimakuMediaInfo {
  if (!mediaPath) {
    return {
      title: '',
      season: null,
      episode: null,
      confidence: 'low',
      filename: '',
      rawTitle: '',
    };
  }

  const guessitInfo = parseMediaInfoWithGuessit(mediaPath);
  if (guessitInfo) return guessitInfo;

  const filename = path.basename(mediaPath);
  let name = filename.replace(/\.[^/.]+$/, '');
  name = name.replace(/\[[^\]]*]/g, ' ');
  name = name.replace(/\(\d{4}\)/g, ' ');
  name = name.replace(/[._]/g, ' ');
  name = name.replace(/[–—]/g, '-');
  name = name.replace(/\s+/g, ' ').trim();

  const parsed = matchEpisodeFromName(name);
  let titlePart = name;
  if (parsed.index !== null) {
    titlePart = name.slice(0, parsed.index);
  }

  const seasonFromDir = parsed.season ?? detectSeasonFromDir(mediaPath);
  const title = cleanupTitle(titlePart || name);

  return {
    title,
    season: seasonFromDir,
    episode: parsed.episode,
    confidence: parsed.confidence,
    filename,
    rawTitle: name,
  };
}

export function sortJimakuFiles(
  files: JimakuFileEntry[],
  pref: JimakuLanguagePreference,
): JimakuFileEntry[] {
  if (pref === 'none') return files;
  return [...files].sort((a, b) => {
    const scoreDiff = formatLangScore(b.name, pref) - formatLangScore(a.name, pref);
    if (scoreDiff !== 0) return scoreDiff;
    return a.name.localeCompare(b.name);
  });
}

export async function downloadToFile(
  url: string,
  destPath: string,
  headers: Record<string, string>,
  redirectCount = 0,
): Promise<JimakuDownloadResult> {
  if (redirectCount > 3) {
    return {
      ok: false,
      error: { error: 'Too many redirects while downloading subtitle.' },
    };
  }

  return new Promise((resolve) => {
    const parsedUrl = new URL(url);
    const transport = parsedUrl.protocol === 'https:' ? https : http;

    const req = transport.get(parsedUrl, { headers }, (res) => {
      const status = res.statusCode || 0;
      if ([301, 302, 303, 307, 308].includes(status) && res.headers.location) {
        const redirectUrl = new URL(res.headers.location, parsedUrl).toString();
        res.resume();
        downloadToFile(redirectUrl, destPath, headers, redirectCount + 1).then(resolve);
        return;
      }

      if (status < 200 || status >= 300) {
        res.resume();
        resolve({
          ok: false,
          error: {
            error: `Failed to download subtitle (HTTP ${status}).`,
            code: status,
          },
        });
        return;
      }

      const fileStream = fs.createWriteStream(destPath);
      res.pipe(fileStream);
      fileStream.on('finish', () => {
        fileStream.close(() => {
          resolve({ ok: true, path: destPath });
        });
      });
      fileStream.on('error', (err: Error) => {
        resolve({
          ok: false,
          error: { error: `Failed to save subtitle: ${err.message}` },
        });
      });
    });

    req.on('error', (err) => {
      resolve({
        ok: false,
        error: {
          error: `Download request failed: ${(err as Error).message}`,
        },
      });
    });
  });
}

export function isValidSubtitleCandidateFile(filename: string): boolean {
  const ext = path.extname(filename).toLowerCase();
  return ext === '.srt' || ext === '.vtt' || ext === '.ass' || ext === '.ssa' || ext === '.sub';
}

export function mapPreferenceToLanguages(preference: JimakuLanguagePreference): string[] {
  if (preference === 'en') return ['en', 'eng'];
  if (preference === 'none') return [];
  return ['ja', 'jpn'];
}

export function normalizeJimakuSearchInput(mediaPath: string): string {
  const trimmed = (mediaPath || '').trim();
  if (!trimmed) return '';
  if (!/^https?:\/\/.*/.test(trimmed)) return trimmed;

  try {
    const url = new URL(trimmed);
    const titleParam =
      url.searchParams.get('title') || url.searchParams.get('name') || url.searchParams.get('q');
    if (titleParam && titleParam.trim()) return titleParam.trim();

    const pathParts = url.pathname.split('/').filter(Boolean).reverse();
    const candidate = pathParts.find((part) => {
      const decoded = decodeURIComponent(part || '').replace(/\.[^/.]+$/, '');
      const lowered = decoded.toLowerCase();
      return lowered.length > 2 && !/^[0-9.]+$/.test(lowered) && !/^[a-f0-9]{16,}$/i.test(lowered);
    });

    const fallback = candidate || url.hostname.replace(/^www\./, '');
    return sanitizeJimakuQueryInput(decodeURIComponent(fallback));
  } catch {
    return trimmed;
  }
}

export function sanitizeJimakuQueryInput(value: string): string {
  return value
    .replace(/^\s*-\s*/, '')
    .replace(/[^\w\s\-'".:(),]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function buildJimakuConfig(args: Args): {
  apiKey: string;
  apiKeyCommand: string;
  apiBaseUrl: string;
  languagePreference: JimakuLanguagePreference;
  maxEntryResults: number;
} {
  return {
    apiKey: args.jimakuApiKey,
    apiKeyCommand: args.jimakuApiKeyCommand,
    apiBaseUrl: args.jimakuApiBaseUrl || DEFAULT_JIMAKU_API_BASE_URL,
    languagePreference: args.jimakuLanguagePreference,
    maxEntryResults: args.jimakuMaxEntryResults || 10,
  };
}
