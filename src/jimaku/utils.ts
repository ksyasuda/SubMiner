import * as http from 'http';
import * as https from 'https';
import * as path from 'path';
import * as fs from 'fs';
import * as childProcess from 'child_process';
import { createLogger } from '../logger';
import { splitSeasonFromTitle } from '../anime-bridge/episode-metadata';
import {
  JimakuApiResponse,
  JimakuConfig,
  JimakuDownloadResult,
  JimakuFileEntry,
  JimakuLanguagePreference,
  JimakuMediaInfo,
} from '../types';

const logger = createLogger('main:jimaku');

function execCommand(command: string): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    childProcess.exec(command, { timeout: 10000 }, (err, stdout, stderr) => {
      if (err) {
        reject(err);
        return;
      }
      resolve({ stdout, stderr });
    });
  });
}

export async function resolveJimakuApiKey(config: JimakuConfig): Promise<string | null> {
  if (config.apiKey && config.apiKey.trim()) {
    logger.debug('API key found in config');
    return config.apiKey.trim();
  }
  if (config.apiKeyCommand && config.apiKeyCommand.trim()) {
    try {
      const { stdout } = await execCommand(config.apiKeyCommand);
      const key = stdout.trim();
      logger.debug(`apiKeyCommand result: ${key.length > 0 ? 'key obtained' : 'empty output'}`);
      return key.length > 0 ? key : null;
    } catch (err) {
      logger.error('Failed to run jimaku.apiKeyCommand', (err as Error).message);
      return null;
    }
  }
  logger.debug('No API key configured (neither apiKey nor apiKeyCommand set)');
  return null;
}

function getRetryAfter(headers: http.IncomingHttpHeaders): number | undefined {
  const value = headers['x-ratelimit-reset-after'];
  if (!value) return undefined;
  const raw = Array.isArray(value) ? value[0] : value;
  const parsed = Number.parseFloat(raw!);
  if (!Number.isFinite(parsed)) return undefined;
  return parsed;
}

export async function jimakuFetchJson<T>(
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined>,
  options: { baseUrl: string; apiKey: string },
): Promise<JimakuApiResponse<T>> {
  const url = new URL(endpoint, options.baseUrl);
  for (const [key, value] of Object.entries(query)) {
    if (value === null || value === undefined) continue;
    url.searchParams.set(key, String(value));
  }

  logger.debug(`GET ${url.toString()}`);
  const transport = url.protocol === 'https:' ? https : http;

  return new Promise((resolve) => {
    const req = transport.request(
      url,
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
          logger.debug(`Response HTTP ${status} for ${endpoint}`);
          if (status >= 200 && status < 300) {
            try {
              const parsed = JSON.parse(data) as T;
              resolve({ ok: true, data: parsed });
            } catch {
              logger.error(`JSON parse error: ${data.slice(0, 200)}`);
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
            // Ignore parse errors.
          }
          logger.error(`API error: ${errorMessage}`);

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

    req.on('error', (err) => {
      logger.error(`Network error: ${(err as Error).message}`);
      resolve({
        ok: false,
        error: { error: `Jimaku request failed: ${(err as Error).message}` },
      });
    });

    req.end();
  });
}

function matchEpisodeFromName(name: string): {
  season: number | null;
  episode: number | null;
  index: number | null;
  confidence: 'high' | 'medium' | 'low';
} {
  const seasonEpisode = name.match(/S(\d{1,2})E(\d{1,3})/i);
  if (seasonEpisode && seasonEpisode.index !== undefined) {
    return {
      season: Number.parseInt(seasonEpisode[1]!, 10),
      episode: Number.parseInt(seasonEpisode[2]!, 10),
      index: seasonEpisode.index,
      confidence: 'high',
    };
  }

  const alt = name.match(/(\d{1,2})x(\d{1,3})/i);
  if (alt && alt.index !== undefined) {
    return {
      season: Number.parseInt(alt[1]!, 10),
      episode: Number.parseInt(alt[2]!, 10),
      index: alt.index,
      confidence: 'high',
    };
  }

  // Spelled-out labels. Streaming sources name episodes this way ("Episode 4"),
  // and the abbreviated `E04` pattern below cannot see them.
  const worded = name.match(/(?:^|[\s._-])episode\s*[.#]?\s*(\d{1,3})(?:\D|$)/i);
  if (worded && worded.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(worded[1]!, 10),
      index: worded.index,
      confidence: 'medium',
    };
  }

  const japanese = name.match(/第\s*(\d{1,3})\s*話/);
  if (japanese && japanese.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(japanese[1]!, 10),
      index: japanese.index,
      confidence: 'medium',
    };
  }

  const epOnly = name.match(/(?:^|[\s._-])E(?:P)?(\d{1,3})(?:\b|[\s._-])/i);
  if (epOnly && epOnly.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(epOnly[1]!, 10),
      index: epOnly.index,
      confidence: 'medium',
    };
  }

  const numeric = name.match(/(?:^|[-–—]\s*)(\d{1,3})\s*[-–—]/);
  if (numeric && numeric.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(numeric[1]!, 10),
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
  const parsed = Number.parseInt(match[1]!, 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function cleanupTitle(value: string): string {
  return value
    .replace(/^[\s-–—]+/, '')
    .replace(/[\s-–—]+$/, '')
    .replace(/\s+/g, ' ')
    .trim();
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

  const normalizedMediaPath = normalizeMediaPathForJimaku(mediaPath);
  const filename = path.basename(normalizedMediaPath);
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

  // A season named in the title ("… Season 3") belongs in the season field, not
  // glued to the title — otherwise it is searched for as part of the name.
  const titleSeason = splitSeasonFromTitle(cleanupTitle(titlePart || name));
  const season = parsed.season ?? titleSeason.season ?? detectSeasonFromDir(normalizedMediaPath);

  return {
    title: titleSeason.title,
    season,
    episode: parsed.episode,
    confidence: parsed.confidence,
    filename,
    rawTitle: name,
  };
}

function normalizeMediaPathForJimaku(mediaPath: string): string {
  const trimmed = mediaPath.trim();
  if (!trimmed || !/^https?:\/\/.*/.test(trimmed)) {
    return trimmed;
  }

  try {
    const parsedUrl = new URL(trimmed);
    const titleParam =
      parsedUrl.searchParams.get('title') ||
      parsedUrl.searchParams.get('name') ||
      parsedUrl.searchParams.get('q');
    if (titleParam && titleParam.trim()) return titleParam.trim();

    const pathParts = parsedUrl.pathname.split('/').filter(Boolean).reverse();
    const candidate = pathParts.find((part) => {
      const decoded = decodeURIComponent(part || '').replace(/\.[^/.]+$/, '');
      const lowered = decoded.toLowerCase();
      return lowered.length > 2 && !/^[0-9.]+$/.test(lowered) && !/^[a-f0-9]{16,}$/i.test(lowered);
    });

    return decodeURIComponent(candidate || parsedUrl.hostname.replace(/^www\./, ''));
  } catch {
    return trimmed;
  }
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

export function isRemoteMediaPath(mediaPath: string): boolean {
  return /^[a-z][a-z0-9+.-]*:\/\//i.test(mediaPath);
}

export function describeDownloadError(err: unknown): string {
  if (err instanceof AggregateError) {
    const parts = err.errors
      .map((inner) => describeDownloadError(inner))
      .filter((part) => part && part !== 'Error');
    if (parts.length > 0) return parts.join('; ');
  }
  if (err instanceof Error) {
    if (err.message) return err.message;
    const code = (err as NodeJS.ErrnoException).code;
    if (code) return code;
    return err.name || 'Error';
  }
  return String(err) || 'Unknown error';
}

export interface DownloadToFileOptions {
  // Guards where a redirect may land. Without it any Location header is followed.
  isAllowedRedirect?: (url: URL) => boolean;
  redirectCount?: number;
}

export async function downloadToFile(
  url: string,
  destPath: string,
  headers: Record<string, string>,
  options: DownloadToFileOptions = {},
): Promise<JimakuDownloadResult> {
  const redirectCount = options.redirectCount ?? 0;
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
        const redirectUrl = new URL(res.headers.location, parsedUrl);
        res.resume();
        if (options.isAllowedRedirect && !options.isAllowedRedirect(redirectUrl)) {
          logger.error(`Refusing redirect to disallowed host: ${redirectUrl.href}`);
          resolve({
            ok: false,
            error: {
              error: `Refusing to follow subtitle redirect to ${redirectUrl.host}.`,
              code: status,
            },
          });
          return;
        }
        downloadToFile(redirectUrl.toString(), destPath, headers, {
          ...options,
          redirectCount: redirectCount + 1,
        }).then(resolve);
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
          error: {
            error: `Failed to save subtitle: ${(err as Error).message}`,
          },
        });
      });
    });

    req.on('error', (err) => {
      const reason = describeDownloadError(err);
      logger.error(`Download request failed for ${url}: ${reason}`);
      resolve({
        ok: false,
        error: { error: `Download request failed: ${reason}` },
      });
    });
  });
}
