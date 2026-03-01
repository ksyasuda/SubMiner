import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { commandExists } from './util.js';

export interface AniSkipMetadata {
  title: string;
  season: number | null;
  episode: number | null;
  source: 'guessit' | 'fallback';
}

interface InferAniSkipDeps {
  commandExists: (name: string) => boolean;
  runGuessit: (mediaPath: string) => string | null;
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
  };
}

function sanitizeScriptOptValue(value: string): string {
  return value
    .replace(/,/g, ' ')
    .replace(/[\r\n]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
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
  return parts.join(',');
}
