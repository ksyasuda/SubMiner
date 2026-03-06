import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { commandExists } from './util.js';

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
  introStart: number | null;
  introEnd: number | null;
  lookupStatus?: AniSkipLookupStatus;
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
  start_time?: unknown;
  end_time?: unknown;
}

interface AniSkipSkipItemPayload {
  skip_type?: unknown;
  interval?: unknown;
}

interface AniSkipPayloadResponse {
  found?: unknown;
  results?: unknown;
}

const MAL_PREFIX_API = 'https://myanimelist.net/search/prefix.json?type=anime&keyword=';
const ANISKIP_PAYLOAD_API = 'https://api.aniskip.com/v1/skip-times/';
const MAL_USER_AGENT = 'SubMiner-launcher/ani-skip';
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

  const romanAliases = {
    2: [' ii ', ' second season ', ' 2nd season '],
    3: [' iii ', ' third season ', ' 3rd season '],
    4: [' iv ', ' fourth season ', ' 4th season '],
    5: [' v ', ' fifth season ', ' 5th season '],
  } as const;

  const aliases = romanAliases[season] ?? [];
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

function normalizeEpisodePayload(value: unknown): number | null {
  return toPositiveNumber(value);
}

function parseAniSkipPayload(payload: unknown): { start: number; end: number } | null {
  const parsed = payload as AniSkipPayloadResponse;
  const results = Array.isArray(parsed?.results) ? parsed.results : null;
  if (!results) return null;

  for (const rawResult of results) {
    const result = rawResult as AniSkipSkipItemPayload;
    if (
      result.skip_type !== 'op' ||
      typeof result.interval !== 'object' ||
      result.interval === null
    ) {
      continue;
    }
    const interval = result.interval as AniSkipIntervalPayload;
    const start = normalizeEpisodePayload(interval?.start_time);
    const end = normalizeEpisodePayload(interval?.end_time);
    if (start !== null && end !== null && end > start) {
      return { start, end };
    }
  }

  return null;
}

async function fetchJson<T>(url: string): Promise<T | null> {
  const response = await fetch(url, {
    headers: {
      'User-Agent': MAL_USER_AGENT,
    },
  });
  if (!response.ok) return null;
  try {
    return (await response.json()) as T;
  } catch {
    return null;
  }
}

async function resolveMalIdFromTitle(title: string, season: number | null): Promise<number | null> {
  const lookup = season && season > 1 ? `${title} Season ${season}` : title;
  const payload = await fetchJson<unknown>(`${MAL_PREFIX_API}${encodeURIComponent(lookup)}`);
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

async function fetchAniSkipPayload(
  malId: number,
  episode: number,
): Promise<{ start: number; end: number } | null> {
  const payload = await fetchJson<unknown>(
    `${ANISKIP_PAYLOAD_API}${malId}/${episode}?types=op&types=ed`,
  );
  const parsed = payload as AniSkipPayloadResponse;
  if (!parsed || parsed.found !== true) return null;
  return parseAniSkipPayload(parsed);
}

function detectEpisodeFromName(baseName: string): number | null {
  const patterns = [
    /[Ss]\d+[Ee](\d{1,3})/,
    /(?:^|[\s._-])[Ee][Pp]?[\s._-]*(\d{1,3})(?:$|[\s._-])/,
    /[-\s](\d{1,3})$/,
  ];
  for (const pattern of patterns) {
    const match = baseName.match(pattern);
    if (!match || !match[1]) continue;
    const parsed = Number.parseInt(match[1], 10);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
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
      introStart: null,
      introEnd: null,
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
  deps: InferAniSkipDeps = { commandExists, runGuessit: defaultRunGuessit },
): AniSkipMetadata {
  if (deps.commandExists('guessit')) {
    const stdout = deps.runGuessit(mediaPath);
    if (typeof stdout === 'string') {
      const parsed = parseAniSkipGuessitJson(stdout, mediaPath);
      if (parsed) return parsed;
    }
  }

  const baseName = path.basename(mediaPath, path.extname(mediaPath));
  const pathTitle = inferTitleFromPath(mediaPath);
  const fallbackTitle = pathTitle || cleanupTitle(baseName) || baseName;
  return {
    title: fallbackTitle,
    season: detectSeasonFromNameOrDir(mediaPath),
    episode: detectEpisodeFromName(baseName),
    source: 'fallback',
    malId: null,
    introStart: null,
    introEnd: null,
    lookupStatus: 'lookup_failed',
  };
}

export async function resolveAniSkipMetadataForFile(mediaPath: string): Promise<AniSkipMetadata> {
  const inferred = inferAniSkipMetadataForFile(mediaPath);
  if (!inferred.title) {
    return { ...inferred, lookupStatus: 'lookup_failed' };
  }

  try {
    const malId = await resolveMalIdFromTitle(inferred.title, inferred.season);
    if (!malId) {
      return {
        ...inferred,
        malId: null,
        introStart: null,
        introEnd: null,
        lookupStatus: 'missing_mal_id',
      };
    }

    if (!inferred.episode) {
      return {
        ...inferred,
        malId,
        introStart: null,
        introEnd: null,
        lookupStatus: 'missing_episode',
      };
    }

    const payload = await fetchAniSkipPayload(malId, inferred.episode);
    if (!payload) {
      return {
        ...inferred,
        malId,
        introStart: null,
        introEnd: null,
        lookupStatus: 'missing_payload',
      };
    }

    return {
      ...inferred,
      malId,
      introStart: payload.start,
      introEnd: payload.end,
      lookupStatus: 'ready',
    };
  } catch {
    return {
      ...inferred,
      malId: inferred.malId,
      introStart: inferred.introStart,
      introEnd: inferred.introEnd,
      lookupStatus: 'lookup_failed',
    };
  }
}

function sanitizeScriptOptValue(value: string): string {
  return value
    .replace(/,/g, ' ')
    .replace(/[\r\n]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildLauncherAniSkipPayload(aniSkipMetadata: AniSkipMetadata): string | null {
  if (!aniSkipMetadata.malId || !aniSkipMetadata.introStart || !aniSkipMetadata.introEnd) {
    return null;
  }
  if (aniSkipMetadata.introEnd <= aniSkipMetadata.introStart) {
    return null;
  }
  const payload = {
    found: true,
    results: [
      {
        skip_type: 'op',
        interval: {
          start_time: aniSkipMetadata.introStart,
          end_time: aniSkipMetadata.introEnd,
        },
      },
    ],
  };
  // mpv --script-opts treats `%` as an escape prefix, so URL-encoding can break parsing.
  // Base64url stays script-opts-safe and is decoded by the plugin launcher payload parser.
  return Buffer.from(JSON.stringify(payload), 'utf-8').toString('base64url');
}

export function buildSubminerScriptOpts(
  appPath: string,
  socketPath: string,
  aniSkipMetadata: AniSkipMetadata | null,
): string {
  const parts = [
    `subminer-binary_path=${sanitizeScriptOptValue(appPath)}`,
    `subminer-socket_path=${sanitizeScriptOptValue(socketPath)}`,
  ];
  if (aniSkipMetadata && aniSkipMetadata.title) {
    parts.push(`subminer-aniskip_title=${sanitizeScriptOptValue(aniSkipMetadata.title)}`);
  }
  if (aniSkipMetadata && aniSkipMetadata.season && aniSkipMetadata.season > 0) {
    parts.push(`subminer-aniskip_season=${aniSkipMetadata.season}`);
  }
  if (aniSkipMetadata && aniSkipMetadata.episode && aniSkipMetadata.episode > 0) {
    parts.push(`subminer-aniskip_episode=${aniSkipMetadata.episode}`);
  }
  if (aniSkipMetadata && aniSkipMetadata.malId && aniSkipMetadata.malId > 0) {
    parts.push(`subminer-aniskip_mal_id=${aniSkipMetadata.malId}`);
  }
  if (aniSkipMetadata && aniSkipMetadata.introStart !== null && aniSkipMetadata.introStart > 0) {
    parts.push(`subminer-aniskip_intro_start=${aniSkipMetadata.introStart}`);
  }
  if (aniSkipMetadata && aniSkipMetadata.introEnd !== null && aniSkipMetadata.introEnd > 0) {
    parts.push(`subminer-aniskip_intro_end=${aniSkipMetadata.introEnd}`);
  }
  if (aniSkipMetadata?.lookupStatus) {
    parts.push(
      `subminer-aniskip_lookup_status=${sanitizeScriptOptValue(aniSkipMetadata.lookupStatus)}`,
    );
  }
  const aniskipPayload = aniSkipMetadata ? buildLauncherAniSkipPayload(aniSkipMetadata) : null;
  if (aniskipPayload) {
    parts.push(`subminer-aniskip_payload=${sanitizeScriptOptValue(aniskipPayload)}`);
  }
  return parts.join(',');
}
