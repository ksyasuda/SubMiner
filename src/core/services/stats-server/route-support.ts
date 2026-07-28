import { existsSync, readFileSync, statSync } from 'node:fs';
import { extname, resolve, sep } from 'node:path';
import {
  getConfiguredSentenceFieldName,
  getConfiguredTranslationFieldName,
  getConfiguredWordFieldName,
  getPreferredNoteFieldValue,
} from '../../../anki-field-config.js';
import {
  knownWordsFromState,
  parseKnownWordCacheState,
} from '../../../anki-integration/known-word-cache-format.js';
import { createLogger } from '../../../logger.js';
import type { AnkiConnectConfig } from '../../../types.js';
import type { StatsExcludedWord } from '../../../types/stats-wire.js';
import type { ImmersionTrackerService } from '../immersion-tracker-service.js';
import { splitSentenceSearchTerms } from '../immersion-tracker/query-lexical.js';

const statsKnownWordsLogger = createLogger('stats:known-words');

const STATS_STATIC_CONTENT_TYPES: Record<string, string> = {
  '.css': 'text/css; charset=utf-8',
  '.gif': 'image/gif',
  '.html': 'text/html; charset=utf-8',
  '.ico': 'image/x-icon',
  '.jpeg': 'image/jpeg',
  '.jpg': 'image/jpeg',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.txt': 'text/plain; charset=utf-8',
  '.webp': 'image/webp',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

export function parseIntQuery(
  raw: string | undefined,
  fallback: number,
  maxLimit?: number,
): number {
  if (raw === undefined) return fallback;
  const n = Number(raw);
  if (!Number.isFinite(n) || n < 0) return fallback;
  const parsed = Math.floor(n);
  return maxLimit === undefined ? parsed : Math.min(parsed, maxLimit);
}

export function parseTrendRange(raw: string | undefined): '7d' | '30d' | '90d' | '365d' | 'all' {
  return raw === '7d' || raw === '30d' || raw === '90d' || raw === '365d' || raw === 'all'
    ? raw
    : '30d';
}

export function parseTrendGroupBy(raw: string | undefined): 'day' | 'month' {
  return raw === 'month' ? 'month' : 'day';
}

export function parseTrendFillEmpty(raw: string | undefined): boolean {
  return raw !== 'false';
}

export function parseEventTypesQuery(raw: string | undefined): number[] | undefined {
  if (!raw) return undefined;
  const parsed = raw
    .split(',')
    .map((entry) => Number.parseInt(entry.trim(), 10))
    .filter((entry) => Number.isInteger(entry) && entry > 0);
  return parsed.length > 0 ? parsed : undefined;
}

export function parseExcludedWordsBody(body: unknown): StatsExcludedWord[] | null {
  if (!body || typeof body !== 'object' || !Array.isArray((body as { words?: unknown }).words)) {
    return null;
  }

  const words: StatsExcludedWord[] = [];
  for (const row of (body as { words: unknown[] }).words) {
    if (!row || typeof row !== 'object') return null;
    const { headword, word, reading } = row as Record<string, unknown>;
    if (typeof headword !== 'string' || typeof word !== 'string' || typeof reading !== 'string') {
      return null;
    }
    words.push({ headword, word, reading });
  }
  return words;
}

export function loadKnownWordsSet(cachePath: string | undefined): Set<string> | null {
  if (!cachePath || !existsSync(cachePath)) return null;
  try {
    const state = parseKnownWordCacheState(JSON.parse(readFileSync(cachePath, 'utf-8')) as unknown);
    if (!state) {
      // A cache that exists but does not parse is a format mismatch, not an
      // empty collection; say so instead of reporting zero known words.
      statsKnownWordsLogger.warn(`Unrecognized known-word cache format at ${cachePath}`);
      return null;
    }
    return knownWordsFromState(state);
  } catch {
    // Treat an unreadable cache as unavailable.
  }
  return null;
}

export function countKnownWords(
  headwords: string[],
  knownWordsSet: Set<string>,
): { totalUniqueWords: number; knownWordCount: number } {
  let knownWordCount = 0;
  for (const headword of headwords) {
    if (knownWordsSet.has(headword)) knownWordCount += 1;
  }
  return { totalUniqueWords: headwords.length, knownWordCount };
}

function toKnownWordRate(knownWordsSeen: number, tokensSeen: number): number {
  if (!Number.isFinite(knownWordsSeen) || !Number.isFinite(tokensSeen) || tokensSeen <= 0) return 0;
  return Number(((knownWordsSeen / tokensSeen) * 100).toFixed(1));
}

function summarizeFilteredWordOccurrences(
  wordsByLine: Array<{ lineIndex: number; headword: string; occurrenceCount: number }>,
  knownWordsSet: Set<string>,
): { knownWordsSeen: number; totalWordsSeen: number } {
  let knownWordsSeen = 0;
  let totalWordsSeen = 0;
  for (const row of wordsByLine) {
    totalWordsSeen += row.occurrenceCount;
    if (knownWordsSet.has(row.headword)) knownWordsSeen += row.occurrenceCount;
  }
  return { knownWordsSeen, totalWordsSeen };
}

export async function enrichSessionsWithKnownWordMetrics<
  Session extends { sessionId: number; tokensSeen: number },
>(
  tracker: ImmersionTrackerService,
  sessions: Session[],
  knownWordsCachePath?: string,
): Promise<Array<Session & { knownWordsSeen: number; knownWordRate: number }>> {
  const knownWordsSet = loadKnownWordsSet(knownWordsCachePath);
  if (!knownWordsSet) {
    return sessions.map((session) => ({ ...session, knownWordsSeen: 0, knownWordRate: 0 }));
  }

  return Promise.all(
    sessions.map(async (session) => {
      let knownWordsSeen = 0;
      let totalWordsSeen = 0;
      try {
        const summary = summarizeFilteredWordOccurrences(
          await tracker.getSessionWordsByLine(session.sessionId),
          knownWordsSet,
        );
        knownWordsSeen = summary.knownWordsSeen;
        totalWordsSeen = summary.totalWordsSeen;
      } catch {
        knownWordsSeen = 0;
        totalWordsSeen = 0;
      }
      return {
        ...session,
        knownWordsSeen,
        knownWordRate: toKnownWordRate(knownWordsSeen, totalWordsSeen),
      };
    }),
  );
}

export function parseBooleanQuery(raw: string | undefined, fallback: boolean): boolean {
  if (raw === undefined) return fallback;
  const normalized = raw.trim().toLowerCase();
  if (!normalized) return fallback;
  return !['0', 'false', 'no', 'off'].includes(normalized);
}

function uniqueNonEmptyStrings(values: readonly string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const value of values) {
    const normalized = value.trim();
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    result.push(normalized);
  }
  return result;
}

export async function buildSentenceSearchOptions(
  query: string,
  searchByHeadword: boolean,
  resolveSentenceSearchHeadwords: ((term: string) => Promise<string[]> | string[]) | undefined,
): Promise<{ headwordTerms: Array<{ term: string; headwords: string[] }> } | undefined> {
  if (!searchByHeadword) return undefined;
  const headwordTerms: Array<{ term: string; headwords: string[] }> = [];
  for (const term of splitSentenceSearchTerms(query)) {
    const resolved = resolveSentenceSearchHeadwords
      ? await resolveSentenceSearchHeadwords(term)
      : [term];
    const headwords = uniqueNonEmptyStrings(resolved);
    if (headwords.length > 0) headwordTerms.push({ term, headwords });
  }
  return headwordTerms.length > 0 ? { headwordTerms } : undefined;
}

export function buildAnkiNotePreview(
  fields: Record<string, { value: string }>,
  ankiConfig?: Pick<AnkiConnectConfig, 'fields'>,
): { word: string; sentence: string; translation: string } {
  return {
    word: getPreferredNoteFieldValue(fields, [getConfiguredWordFieldName(ankiConfig)]),
    sentence: getPreferredNoteFieldValue(fields, [getConfiguredSentenceFieldName(ankiConfig)]),
    translation: getPreferredNoteFieldValue(fields, [
      getConfiguredTranslationFieldName(ankiConfig),
    ]),
  };
}

function resolveStatsStaticPath(staticDir: string, requestPath: string): string | null {
  const normalizedPath = requestPath.replace(/^\/+/, '') || 'index.html';
  const absoluteStaticDir = resolve(staticDir);
  let decodedPath: string;
  try {
    decodedPath = decodeURIComponent(normalizedPath);
  } catch {
    return null;
  }
  const absolutePath = resolve(absoluteStaticDir, decodedPath);
  if (
    absolutePath !== absoluteStaticDir &&
    !absolutePath.startsWith(`${absoluteStaticDir}${sep}`)
  ) {
    return null;
  }
  if (!existsSync(absolutePath) || !statSync(absolutePath).isFile()) return null;
  return absolutePath;
}

export function createStatsStaticResponse(staticDir: string, requestPath: string): Response | null {
  const absolutePath = resolveStatsStaticPath(staticDir, requestPath);
  if (!absolutePath) return null;
  const contentType =
    STATS_STATIC_CONTENT_TYPES[extname(absolutePath).toLowerCase()] ?? 'application/octet-stream';
  return new Response(readFileSync(absolutePath), {
    headers: {
      'Content-Type': contentType,
      'Cache-Control': absolutePath.endsWith('index.html')
        ? 'no-cache'
        : 'public, max-age=31536000, immutable',
    },
  });
}
