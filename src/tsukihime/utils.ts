import * as http from 'http';
import * as https from 'https';
import * as fs from 'fs';
import * as path from 'path';
import * as childProcess from 'child_process';
import { createLogger } from '../logger';
import {
  TsukihimeApiResponse,
  TsukihimeDownloadResult,
  TsukihimeEntry,
  TsukihimeSubtitleFile,
} from '../types';

// Backend: TsukiHime (https://tsukihime.org), the community successor to
// Animetosho (read-only from May 2026). TsukiHime imported the Animetosho
// index, so "animetosho" still appears here where it names the upstream
// service: the imported-entry flag, the /tosho/ storage mirror, and the
// redirect host that mirror currently points at.
const logger = createLogger('main:tsukihime');

export const TSUKIHIME_API_BASE_URL = 'https://api.tsukihime.org/v1';
export const TSUKIHIME_STORAGE_BASE_URL = 'https://storage.tsukihime.org';

// Entries imported from the Animetosho index sit in this id range and their
// attachments live under the /tosho/ mirror path on the storage host.
const IMPORTED_ANIMETOSHO_ID_FLOOR = 1_000_000_000;

const TEXT_SUBTITLE_EXTENSIONS: Record<string, string> = {
  ass: '.ass',
  ssa: '.ssa',
  srt: '.srt',
  subrip: '.srt',
};

const XZ_MAX_SUBTITLE_BYTES = 64 * 1024 * 1024;
const TSUKIHIME_REQUEST_TIMEOUT_MS = 15000;
const TSUKIHIME_MAX_RESPONSE_BYTES = 16 * 1024 * 1024;

export {
  tsukihimeLangToFilenameSuffix,
  tsukihimeTrackMatchesLanguages,
  describeTsukihimeTabLanguages,
  normalizeTsukihimeLangCode,
} from './lang';

export function isTsukihimeDownloadUrl(url: string | URL): boolean {
  try {
    const parsed = typeof url === 'string' ? new URL(url) : url;
    if (parsed.protocol !== 'https:') return false;
    // The /tosho/ mirror currently 302s to storage.animetosho.org, so both
    // hosts must stay allowed for the redirect hop.
    const allowedHosts = ['tsukihime.org', 'animetosho.org'];
    return allowedHosts.some(
      (host) => parsed.hostname === host || parsed.hostname.endsWith(`.${host}`),
    );
  } catch {
    return false;
  }
}

export function buildTsukihimeAttachmentUrl(
  attachmentId: number,
  options?: { imported?: boolean },
): string | null {
  if (!Number.isInteger(attachmentId) || attachmentId <= 0) return null;
  const hexId = attachmentId.toString(16).padStart(8, '0');
  const attachPath = options?.imported ? '/tosho/attach/' : '/attach/';
  return `${TSUKIHIME_STORAGE_BASE_URL}${attachPath}${hexId}/${attachmentId}.xz`;
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function asFiniteNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

export function mapTsukihimeSearchResults(
  payload: unknown,
  maxResults: number,
): TsukihimeEntry[] {
  if (!isObject(payload) || !Array.isArray(payload.results)) return [];
  const entries: TsukihimeEntry[] = [];
  for (const item of payload.results) {
    if (entries.length >= maxResults) break;
    if (!isObject(item)) continue;
    const id = item.id;
    const title = item.name;
    if (!Number.isInteger(id) || typeof title !== 'string' || !title) continue;
    const sublangs = Array.isArray(item.sublangs)
      ? item.sublangs.filter((lang): lang is string => typeof lang === 'string')
      : [];
    entries.push({
      id: id as number,
      title,
      timestamp: asFiniteNumber(item.added_date),
      totalSize: asFiniteNumber(item.totalsize),
      numFiles: asFiniteNumber(item.filecount),
      sublangs,
    });
  }
  return entries;
}

function stripFileExtension(name: string): string {
  const ext = path.extname(name);
  return ext ? name.slice(0, -ext.length) : name;
}

function slugifyTrackName(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function subtitleLangRank(lang: string): number {
  if (lang.startsWith('en')) return 0;
  if (!lang || lang === 'und') return 1;
  return 2;
}

interface RawSubtitleAttachment {
  attachmentId: number;
  lang: string;
  trackName: string | null;
  size: number;
  url: string;
  sourceFilename: string;
  extension: string;
}

const SUBTITLE_ATTACHMENT_TYPE = 1;

function isImportedAnimetoshoEntry(payload: Record<string, unknown>): boolean {
  if (payload.animetosho === true) return true;
  const id = asFiniteNumber(payload.id);
  return id !== null && id >= IMPORTED_ANIMETOSHO_ID_FLOOR;
}

function collectSubtitleAttachments(payload: unknown): RawSubtitleAttachment[] {
  if (!isObject(payload) || !Array.isArray(payload.files)) return [];
  const imported = isImportedAnimetoshoEntry(payload);
  const collected: RawSubtitleAttachment[] = [];
  for (const file of payload.files) {
    if (!isObject(file) || !Array.isArray(file.attachments)) continue;
    const sourceFilename = typeof file.filename === 'string' && file.filename ? file.filename : '';
    for (const attachment of file.attachments) {
      if (!isObject(attachment) || attachment.type !== SUBTITLE_ATTACHMENT_TYPE) continue;
      const attachmentId = attachment.id;
      if (!Number.isInteger(attachmentId) || (attachmentId as number) <= 0) continue;
      const info = isObject(attachment.info) ? attachment.info : {};
      const codec = typeof info.codec === 'string' ? info.codec.toLowerCase() : '';
      const extension = TEXT_SUBTITLE_EXTENSIONS[codec];
      if (!extension) continue;
      const url = buildTsukihimeAttachmentUrl(attachmentId as number, { imported });
      if (!url) continue;
      collected.push({
        attachmentId: attachmentId as number,
        lang: typeof info.lang === 'string' ? info.lang.toLowerCase() : '',
        trackName: typeof info.name === 'string' && info.name ? info.name : null,
        size: asFiniteNumber(attachment.size) ?? 0,
        url,
        sourceFilename,
        extension,
      });
    }
  }
  return collected;
}

export function extractTsukihimeSubtitleFiles(payload: unknown): TsukihimeSubtitleFile[] {
  const collected = collectSubtitleAttachments(payload);

  const duplicateKeyCounts = new Map<string, number>();
  for (const attachment of collected) {
    const key = `${attachment.sourceFilename} ${attachment.lang}`;
    duplicateKeyCounts.set(key, (duplicateKeyCounts.get(key) ?? 0) + 1);
  }

  const files = collected.map((attachment) => {
    const base = attachment.sourceFilename
      ? stripFileExtension(attachment.sourceFilename)
      : 'subtitle';
    const langPart = attachment.lang ? `.${attachment.lang}` : '';
    const key = `${attachment.sourceFilename} ${attachment.lang}`;
    const needsTrackSuffix = (duplicateKeyCounts.get(key) ?? 0) > 1 && attachment.trackName;
    const trackPart = needsTrackSuffix ? `.${slugifyTrackName(attachment.trackName!)}` : '';
    return {
      attachmentId: attachment.attachmentId,
      filename: `${base}${langPart}${trackPart}${attachment.extension}`,
      lang: attachment.lang,
      trackName: attachment.trackName,
      size: attachment.size,
      url: attachment.url,
      sourceFilename: attachment.sourceFilename,
    };
  });

  return files
    .map((file, index) => ({ file, index }))
    .sort((a, b) => {
      const rankDiff = subtitleLangRank(a.file.lang) - subtitleLangRank(b.file.lang);
      if (rankDiff !== 0) return rankDiff;
      return a.index - b.index;
    })
    .map(({ file }) => file);
}

export async function tsukihimeFetchJson<T>(
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined>,
  options: { baseUrl: string; timeoutMs?: number; maxResponseBytes?: number },
): Promise<TsukihimeApiResponse<T>> {
  // Concatenate instead of new URL(endpoint, base): the API base ends in a
  // path segment (/v1) that URL resolution would drop for absolute endpoints.
  const url = new URL(`${options.baseUrl.replace(/\/+$/, '')}${endpoint}`);
  for (const [key, value] of Object.entries(query)) {
    if (value === null || value === undefined) continue;
    url.searchParams.set(key, String(value));
  }

  const timeoutMs = options.timeoutMs ?? TSUKIHIME_REQUEST_TIMEOUT_MS;
  const maxResponseBytes = options.maxResponseBytes ?? TSUKIHIME_MAX_RESPONSE_BYTES;

  logger.debug(`GET ${url.toString()}`);
  const transport = url.protocol === 'https:' ? https : http;

  return new Promise((resolve) => {
    let settled = false;
    let timeoutHandle: ReturnType<typeof setTimeout> | null = null;
    const settle = (result: TsukihimeApiResponse<T>): void => {
      if (settled) return;
      settled = true;
      if (timeoutHandle) clearTimeout(timeoutHandle);
      resolve(result);
    };

    const req = transport.request(
      url,
      {
        method: 'GET',
        headers: { 'User-Agent': 'SubMiner' },
      },
      (res) => {
        // Buffer raw chunks and decode once: a multi-byte UTF-8 character
        // (e.g. a Japanese title) can straddle a chunk boundary.
        const chunks: Buffer[] = [];
        let receivedBytes = 0;
        res.on('data', (chunk: Buffer) => {
          receivedBytes += chunk.length;
          if (receivedBytes > maxResponseBytes) {
            logger.error(`TsukiHime response exceeded ${maxResponseBytes} bytes; aborting.`);
            // Settle before destroying: destroy can emit 'end' synchronously.
            settle({
              ok: false,
              error: { error: 'TsukiHime response was too large.' },
            });
            req.destroy();
            return;
          }
          chunks.push(chunk);
        });
        res.on('end', () => {
          if (settled) return;
          const status = res.statusCode || 0;
          logger.debug(`Response HTTP ${status} for ${endpoint}`);
          if (status >= 200 && status < 300) {
            const data = Buffer.concat(chunks).toString('utf8');
            try {
              const parsed = JSON.parse(data) as T;
              settle({ ok: true, data: parsed });
            } catch {
              logger.error(`JSON parse error: ${data.slice(0, 200)}`);
              settle({
                ok: false,
                error: { error: 'Failed to parse TsukiHime response JSON.' },
              });
            }
            return;
          }
          logger.error(`TsukiHime API error (HTTP ${status})`);
          settle({
            ok: false,
            error: { error: `TsukiHime API error (HTTP ${status})`, code: status || undefined },
          });
        });
      },
    );

    timeoutHandle = setTimeout(() => {
      logger.error(`TsukiHime request timed out after ${timeoutMs}ms: ${url.toString()}`);
      settle({
        ok: false,
        error: { error: `TsukiHime request timed out after ${timeoutMs}ms.` },
      });
      req.destroy();
    }, timeoutMs);

    req.on('error', (err) => {
      logger.error(`Network error: ${(err as Error).message}`);
      settle({
        ok: false,
        error: { error: `TsukiHime request failed: ${(err as Error).message}` },
      });
    });

    req.end();
  });
}

export async function decompressXzFile(
  srcPath: string,
  destPath: string,
): Promise<TsukihimeDownloadResult> {
  return new Promise((resolve) => {
    childProcess.execFile(
      'xz',
      ['-dc', srcPath],
      { encoding: 'buffer', maxBuffer: XZ_MAX_SUBTITLE_BYTES },
      async (err, stdout) => {
        if (err) {
          const reason =
            (err as NodeJS.ErrnoException).code === 'ENOENT'
              ? 'xz binary not found. Install xz-utils to download TsukiHime subtitles.'
              : `Failed to decompress subtitle with xz: ${err.message}`;
          logger.error(reason);
          resolve({ ok: false, error: { error: reason } });
          return;
        }
        try {
          await fs.promises.writeFile(destPath, stdout);
          resolve({ ok: true, path: destPath });
        } catch (writeErr) {
          resolve({
            ok: false,
            error: { error: `Failed to save subtitle: ${(writeErr as Error).message}` },
          });
        }
      },
    );
  });
}
