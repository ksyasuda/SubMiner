import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import type { SkipInterval } from './aniskip-chapters';

export type AniSkipLookupStatus =
  | 'ready'
  | 'missing_mal_id'
  | 'missing_episode'
  | 'missing_payload'
  | 'lookup_failed';

export interface AniSkipMetadata {
  title: string;
  season: number | null;
  episode: number | null;
  source: 'guessit' | 'fallback';
  malId: number | null;
  intro: SkipInterval | null;
  outro: SkipInterval | null;
  lookupStatus?: AniSkipLookupStatus;
}

export interface AniSkipLookupOptions {
  /** Length of the playing file in seconds; AniSkip v2 uses it to pick times marked on the same cut. */
  episodeLength: number;
}

interface InferAniSkipDeps {
  commandExists: (name: string) => boolean;
  runGuessit: (mediaPath: string) => string | null;
}

interface MalSearchResult {
  id?: unknown;
  name?: unknown;
}

interface MalSearchCategory {
  items?: unknown;
}

interface MalSearchResponse {
  categories?: unknown;
}

interface AniSkipIntervalPayload {
  startTime?: unknown;
  endTime?: unknown;
}

interface AniSkipSkipItemPayload {
  skipType?: unknown;
  interval?: unknown;
}

interface AniSkipPayloadResponse {
  found?: unknown;
  results?: unknown;
}

const ROMAN_SEASON_ALIASES: Record<number, readonly string[]> = {
  2: [' ii ', ' second season ', ' 2nd season '],
  3: [' iii ', ' third season ', ' 3rd season '],
  4: [' iv ', ' fourth season ', ' 4th season '],
  5: [' v ', ' fifth season ', ' 5th season '],
};

const MAL_PREFIX_API = 'https://myanimelist.net/search/prefix.json?type=anime&keyword=';
const ANISKIP_PAYLOAD_API = 'https://api.aniskip.com/v2/skip-times/';
const ANISKIP_USER_AGENT = 'SubMiner/ani-skip';
const ANISKIP_FETCH_TIMEOUT_MS = 8000;
// AniSkip allows 120 requests a minute; one lookup per file load stays far below that,
// but a gap keeps fast playlist skipping from bursting.
const ANISKIP_MIN_REQUEST_GAP_MS = 1000;
const MAL_MATCH_STOPWORDS = new Set([
  'the',
  'this',
  'that',
  'world',
  'animated',
  'series',
  'season',
  'no',
  'on',
  'and',
]);

function commandExistsOnPath(command: string): boolean {
  const pathEnv = process.env.PATH || '';
  const extensions =
    process.platform === 'win32' ? (process.env.PATHEXT || '.EXE;.CMD;.BAT;.COM').split(';') : [''];
  for (const dir of pathEnv.split(path.delimiter)) {
    if (!dir) continue;
    for (const extension of extensions) {
      const candidate = path.join(dir, `${command}${extension.toLowerCase()}`);
      try {
        fs.accessSync(candidate, fs.constants.X_OK);
        if (fs.statSync(candidate).isFile()) {
          return true;
        }
      } catch {
        // keep scanning PATH entries
      }
    }
  }
  return false;
}

function toPositiveInt(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value) && value > 0) {
    return Math.floor(value);
  }
  if (typeof value === 'string') {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed) && parsed > 0) {
      return parsed;
    }
  }
  return null;
}

function toPositiveNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value) && value > 0) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number.parseFloat(value);
    if (Number.isFinite(parsed) && parsed > 0) {
      return parsed;
    }
  }
  return null;
}

function toNonNegativeNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value) && value >= 0) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number.parseFloat(value);
    if (Number.isFinite(parsed) && parsed >= 0) {
      return parsed;
    }
  }
  return null;
}

function normalizeForMatch(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^\w]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tokenizeMatchWords(value: string): string[] {
  const words = normalizeForMatch(value)
    .split(' ')
    .filter((word) => word.length >= 3);
  return words.filter((word) => !MAL_MATCH_STOPWORDS.has(word));
}

function titleOverlapScore(expectedTitle: string, candidateTitle: string): number {
  const expected = normalizeForMatch(expectedTitle);
  const candidate = normalizeForMatch(candidateTitle);

  if (!expected || !candidate) return 0;

  if (candidate.includes(expected)) return 120;
  if (candidate.split(' ').length >= 2 && ` ${expected} `.includes(` ${candidate} `)) {
    return 90;
  }

  const expectedTokens = tokenizeMatchWords(expectedTitle);
  if (expectedTokens.length === 0) return 0;

  const candidateSet = new Set(tokenizeMatchWords(candidateTitle));
  let score = 0;
  let matched = 0;

  for (const token of expectedTokens) {
    if (candidateSet.has(token)) {
      score += 30;
      matched += 1;
    } else {
      score -= 20;
    }
  }

  if (matched === 0) {
    score -= 80;
  }

  const coverage = matched / expectedTokens.length;
  if (expectedTokens.length >= 2) {
    if (coverage >= 0.8) score += 30;
    else if (coverage >= 0.6) score += 10;
    else score -= 50;
  } else if (coverage >= 1) {
    score += 10;
  }

  return score;
}

function hasAnySequelMarker(candidateTitle: string): boolean {
  const normalized = ` ${normalizeForMatch(candidateTitle)} `;
  if (!normalized.trim()) return false;

  const markers = [
    'season 2',
    'season 3',
    'season 4',
    '2nd season',
    '3rd season',
    '4th season',
    'second season',
    'third season',
    'fourth season',
    ' ii ',
    ' iii ',
    ' iv ',
  ];
  return markers.some((marker) => normalized.includes(marker));
}

function seasonSignalScore(requestedSeason: number | null, candidateTitle: string): number {
  const season = toPositiveInt(requestedSeason);
  if (!season || season < 1) return 0;

  const normalized = ` ${normalizeForMatch(candidateTitle)} `;
  if (!normalized.trim()) return 0;

  if (season === 1) {
    return hasAnySequelMarker(candidateTitle) ? -60 : 20;
  }

  const numericMarker = ` season ${season} `;
  const ordinalMarker = ` ${season}th season `;
  if (normalized.includes(numericMarker) || normalized.includes(ordinalMarker)) {
    return 40;
  }

  const aliases = ROMAN_SEASON_ALIASES[season] ?? [];
  return aliases.some((alias) => normalized.includes(alias))
    ? 40
    : hasAnySequelMarker(candidateTitle)
      ? -20
      : 5;
}

function toMalSearchItems(payload: unknown): MalSearchResult[] {
  const parsed = payload as MalSearchResponse;
  const categories = Array.isArray(parsed?.categories) ? parsed.categories : null;
  if (!categories) return [];

  const items: MalSearchResult[] = [];
  for (const category of categories) {
    const typedCategory = category as MalSearchCategory;
    const rawItems = Array.isArray(typedCategory?.items) ? typedCategory.items : [];
    for (const rawItem of rawItems) {
      const item = rawItem as Record<string, unknown>;
      items.push({
        id: item?.id,
        name: item?.name,
      });
    }
  }
  return items;
}

function parseInterval(raw: unknown): SkipInterval | null {
  if (typeof raw !== 'object' || raw === null) return null;
  const interval = raw as AniSkipIntervalPayload;
  const start = toNonNegativeNumber(interval.startTime);
  const end = toPositiveNumber(interval.endTime);
  return start !== null && end !== null && end > start ? { start, end } : null;
}

/** Reads the first valid opening (`op`) and ending (`ed`) from an AniSkip v2 response. */
export function parseAniSkipPayload(payload: unknown): {
  intro: SkipInterval | null;
  outro: SkipInterval | null;
} {
  const parsed = payload as AniSkipPayloadResponse | null;
  const results = parsed?.found === true && Array.isArray(parsed.results) ? parsed.results : [];
  let intro: SkipInterval | null = null;
  let outro: SkipInterval | null = null;
  for (const rawResult of results) {
    const result = rawResult as AniSkipSkipItemPayload;
    if (result.skipType === 'op' && !intro) intro = parseInterval(result.interval);
    if (result.skipType === 'ed' && !outro) outro = parseInterval(result.interval);
  }
  return { intro, outro };
}

/**
 * Fetches JSON. A 404 is a definite "nothing here" and returns null; any other
 * failure (non-OK status, network error, timeout, bad JSON) throws so callers
 * report a retryable lookup failure instead of caching a miss.
 */
async function fetchJson(url: string): Promise<unknown> {
  const response = await fetch(url, {
    headers: {
      'User-Agent': ANISKIP_USER_AGENT,
    },
    signal: AbortSignal.timeout(ANISKIP_FETCH_TIMEOUT_MS),
  });
  if (response.status === 404) return null;
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} from ${new URL(url).host}`);
  }
  return await response.json();
}

let lastAniSkipRequestAt = 0;
let aniSkipRequestQueue: Promise<void> = Promise.resolve();

/**
 * Resolves when the caller may send its AniSkip request. Callers are released
 * one at a time, each at least ANISKIP_MIN_REQUEST_GAP_MS after the previous
 * release, so timers that fire late cannot release a burst.
 */
function waitForAniSkipRequestSlot(): Promise<void> {
  const slot = aniSkipRequestQueue.then(async () => {
    const waitMs = lastAniSkipRequestAt + ANISKIP_MIN_REQUEST_GAP_MS - Date.now();
    if (waitMs > 0) await new Promise((resolve) => setTimeout(resolve, waitMs));
    lastAniSkipRequestAt = Date.now();
  });
  aniSkipRequestQueue = slot;
  return slot;
}

async function resolveMalIdFromTitle(title: string, season: number | null): Promise<number | null> {
  const lookup = season && season > 1 ? `${title} Season ${season}` : title;
  const payload = await fetchJson(`${MAL_PREFIX_API}${encodeURIComponent(lookup)}`);
  const items = toMalSearchItems(payload);
  if (!items.length) return null;

  let bestScore = Number.NEGATIVE_INFINITY;
  let bestMalId: number | null = null;

  for (const item of items) {
    const id = toPositiveInt(item.id);
    if (!id) continue;
    const name = typeof item.name === 'string' ? item.name : '';
    if (!name) continue;

    const score = titleOverlapScore(title, name) + seasonSignalScore(season, name);
    if (score > bestScore) {
      bestScore = score;
      bestMalId = id;
    }
  }

  return bestMalId;
}

async function fetchAniSkipIntervals(
  malId: number,
  episode: number,
  episodeLength: number,
): Promise<{ intro: SkipInterval | null; outro: SkipInterval | null }> {
  await waitForAniSkipRequestSlot();
  const length = Math.round(episodeLength * 1000) / 1000;
  const payload = await fetchJson(
    `${ANISKIP_PAYLOAD_API}${malId}/${episode}?types=op&types=ed&episodeLength=${length}`,
  );
  return parseAniSkipPayload(payload);
}

function detectEpisodeFromName(baseName: string): number | null {
  const patterns = [
    /[Ss]\d+[Ee](\d{1,3})/,
    /(?:^|[\s._-])[Ee][Pp]?[\s._-]*(\d{1,3})(?:$|[\s._-])/,
    /(?:^|[\s._-])(\d{1,3})(?:$|[\s._-])/,
    /[-\s](\d{1,3})$/,
  ];
  const groupStrippedName = baseName.replace(/\[[^\]]+\]/g, ' ').replace(/\([^)]+\)/g, ' ');
  for (const candidate of [baseName, groupStrippedName]) {
    for (const pattern of patterns) {
      const match = candidate.match(pattern);
      if (!match || !match[1]) continue;
      const parsed = Number.parseInt(match[1], 10);
      if (Number.isFinite(parsed) && parsed > 0) return parsed;
    }
  }
  return null;
}

function detectSeasonFromNameOrDir(mediaPath: string): number | null {
  const baseName = path.basename(mediaPath, path.extname(mediaPath));
  const seasonMatch = baseName.match(/[Ss](\d{1,2})[Ee]\d{1,3}/);
  if (seasonMatch && seasonMatch[1]) {
    const parsed = Number.parseInt(seasonMatch[1], 10);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
  }
  const parent = path.basename(path.dirname(mediaPath));
  const parentMatch = parent.match(/(?:Season|S)[\s._-]*(\d{1,2})/i);
  if (parentMatch && parentMatch[1]) {
    const parsed = Number.parseInt(parentMatch[1], 10);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
  }
  return null;
}

function isSeasonDirectoryName(value: string): boolean {
  return /^(?:season|s)[\s._-]*\d{1,2}$/i.test(value.trim());
}

function isEpisodeOnlyBaseName(value: string): boolean {
  return /^(?:[Ss]\d{1,2}[Ee]\d{1,3}|[Ee][Pp]?[\s._-]*\d{1,3}|\d{1,3})(?:$|[\s._-])/.test(
    value.trim(),
  );
}

function inferTitleFromPath(mediaPath: string): string {
  const directory = path.dirname(mediaPath);
  const segments = directory.split(/[\\/]+/).filter((segment) => segment.length > 0);
  for (let index = 0; index < segments.length; index += 1) {
    const segment = segments[index] || '';
    if (!isSeasonDirectoryName(segment)) continue;
    const showSegment = segments[index - 1];
    if (typeof showSegment === 'string' && showSegment.length > 0) {
      const cleaned = cleanupTitle(showSegment);
      if (cleaned) return cleaned;
    }
  }

  const parent = path.basename(directory);
  if (!isSeasonDirectoryName(parent)) {
    const cleanedParent = cleanupTitle(parent);
    if (cleanedParent) return cleanedParent;
  }

  const grandParent = path.basename(path.dirname(directory));
  const cleanedGrandParent = cleanupTitle(grandParent);
  return cleanedGrandParent;
}

function cleanupTitle(value: string): string {
  return value
    .replace(/\.[^/.]+$/, '')
    .replace(/\[[^\]]+\]/g, ' ')
    .replace(/\([^)]+\)/g, ' ')
    .replace(/[Ss]\d+[Ee]\d+/g, ' ')
    .replace(/[Ee][Pp]?[\s._-]*\d+/g, ' ')
    .replace(/(?:^|[\s._-])\d{1,3}\s*$/g, ' ')
    .replace(/[_\-.]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function parseAniSkipGuessitJson(stdout: string, mediaPath: string): AniSkipMetadata | null {
  const payload = stdout.trim();
  if (!payload) return null;

  try {
    const parsed = JSON.parse(payload) as {
      title?: unknown;
      title_original?: unknown;
      series?: unknown;
      season?: unknown;
      episode?: unknown;
      episode_list?: unknown;
    };

    const rawTitle =
      (typeof parsed.series === 'string' && parsed.series) ||
      (typeof parsed.title === 'string' && parsed.title) ||
      (typeof parsed.title_original === 'string' && parsed.title_original) ||
      '';

    const title = cleanupTitle(rawTitle) || inferTitleFromPath(mediaPath);
    if (!title) return null;

    const season = toPositiveInt(parsed.season);
    const episodeFromDirect = toPositiveInt(parsed.episode);
    const episodeFromList =
      Array.isArray(parsed.episode_list) && parsed.episode_list.length > 0
        ? toPositiveInt(parsed.episode_list[0])
        : null;

    return {
      title,
      season,
      episode: episodeFromDirect ?? episodeFromList,
      source: 'guessit',
      malId: null,
      intro: null,
      outro: null,
      lookupStatus: 'lookup_failed',
    };
  } catch {
    return null;
  }
}

function defaultRunGuessit(mediaPath: string): string | null {
  const fileName = path.basename(mediaPath);
  const result = spawnSync('guessit', ['--json', fileName], {
    cwd: path.dirname(mediaPath),
    encoding: 'utf8',
    maxBuffer: 2_000_000,
    windowsHide: true,
  });
  if (result.error || result.status !== 0) return null;
  return result.stdout || null;
}

export function inferAniSkipMetadataForFile(
  mediaPath: string,
  deps: InferAniSkipDeps = { commandExists: commandExistsOnPath, runGuessit: defaultRunGuessit },
): AniSkipMetadata {
  if (deps.commandExists('guessit')) {
    const stdout = deps.runGuessit(mediaPath);
    if (typeof stdout === 'string') {
      const parsed = parseAniSkipGuessitJson(stdout, mediaPath);
      if (parsed) return parsed;
    }
  }

  const baseName = path.basename(mediaPath, path.extname(mediaPath));
  const cleanedBaseName = cleanupTitle(baseName);
  const pathTitle = inferTitleFromPath(mediaPath);
  const fallbackTitle = isEpisodeOnlyBaseName(baseName)
    ? pathTitle || cleanedBaseName || baseName
    : cleanedBaseName || pathTitle || baseName;
  return {
    title: fallbackTitle,
    season: detectSeasonFromNameOrDir(mediaPath),
    episode: detectEpisodeFromName(baseName),
    source: 'fallback',
    malId: null,
    intro: null,
    outro: null,
    lookupStatus: 'lookup_failed',
  };
}

/**
 * Resolves a local file's MAL id and AniSkip opening/ending. Statuses other
 * than `lookup_failed` are definite answers that callers may cache.
 */
export async function resolveAniSkipMetadataForFile(
  mediaPath: string,
  options: AniSkipLookupOptions,
): Promise<AniSkipMetadata> {
  const inferred = inferAniSkipMetadataForFile(mediaPath);
  if (!inferred.title) {
    return { ...inferred, lookupStatus: 'lookup_failed' };
  }

  try {
    const malId = await resolveMalIdFromTitle(inferred.title, inferred.season);
    if (!malId) {
      return { ...inferred, lookupStatus: 'missing_mal_id' };
    }
    if (!inferred.episode) {
      return { ...inferred, malId, lookupStatus: 'missing_episode' };
    }

    const { intro, outro } = await fetchAniSkipIntervals(
      malId,
      inferred.episode,
      options.episodeLength,
    );
    return {
      ...inferred,
      malId,
      intro,
      outro,
      lookupStatus: intro || outro ? 'ready' : 'missing_payload',
    };
  } catch {
    return { ...inferred, lookupStatus: 'lookup_failed' };
  }
}
