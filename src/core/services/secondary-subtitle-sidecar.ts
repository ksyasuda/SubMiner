import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { runCommand, type CommandResult } from '../../subsync/utils';
import { resolveExecutable, SUBSYNC_EXECUTABLE_NAMES } from '../../subsync/executables';
import { parseSubtitleCues, type SubtitleCue } from './subtitle-cue-parser.js';
import { isEnglishYoutubeLang, normalizeYoutubeLangCode } from './youtube/labels.js';

const DEFAULT_SECONDARY_SUBTITLE_LANGUAGES = ['en', 'eng', 'english', 'en-us', 'enus'];
const DEFAULT_PRIMARY_SUBTITLE_LANGUAGES = ['ja', 'jpn', 'jp', 'japanese'];
const SUPPORTED_SUBTITLE_EXTENSIONS = new Set(['.srt', '.vtt', '.ass', '.ssa']);
const TIMING_TOLERANCE_SECONDS = 0.25;
const SAME_TIMING_EPSILON_SECONDS = 0.001;
const RETIMED_SUBTITLE_TIMEOUT_MS = 30_000;

type SidecarCandidate = {
  path: string;
  languageRank: number;
  extensionRank: number;
  name: string;
};

type RetimedSubtitleCacheEntry = {
  path: string;
  cleanupDir: string;
  promise?: Promise<string>;
};

export type RetimedSubtitleCommandRunner = (
  alassPath: string,
  referencePath: string,
  inputPath: string,
  outputPath: string,
) => Promise<CommandResult>;

export type RetimedSecondarySubtitleInput = {
  sourcePath: string;
  startMs: number;
  endMs: number;
  languages?: readonly string[];
  primaryLanguages?: readonly string[];
  alassPath?: string | null;
  runAlass?: RetimedSubtitleCommandRunner;
};

const retimedSubtitleCache = new Map<string, RetimedSubtitleCacheEntry>();
let retimedSubtitleCleanupRegistered = false;

function unique(values: string[]): string[] {
  return values.filter((value, index) => value.length > 0 && values.indexOf(value) === index);
}

function expandPreferredLanguages(
  languages: readonly string[] | undefined,
  fallback: readonly string[],
): string[] {
  const normalized = unique(
    (languages ?? []).map((language) => normalizeYoutubeLangCode(language)).filter(Boolean),
  );
  const base = normalized.length > 0 ? normalized : [...fallback];
  const expanded: string[] = [];
  for (const language of base) {
    expanded.push(language);
    if (isEnglishYoutubeLang(language)) {
      expanded.push(...DEFAULT_SECONDARY_SUBTITLE_LANGUAGES);
    }
  }
  return unique(expanded);
}

function resolveAlassPath(configuredPath: string | null | undefined): string {
  return resolveExecutable(configuredPath, SUBSYNC_EXECUTABLE_NAMES.alass);
}

function fileSignature(filePath: string): string | null {
  try {
    const stats = statSync(filePath);
    if (!stats.isFile()) return null;
    return `${stats.size}:${stats.mtimeMs}`;
  } catch {
    return null;
  }
}

function retimedCacheKey(
  alassPath: string,
  primaryPath: string,
  secondaryPath: string,
): string | null {
  const primarySignature = fileSignature(primaryPath);
  const secondarySignature = fileSignature(secondaryPath);
  if (!primarySignature || !secondarySignature) return null;
  return [alassPath, primaryPath, primarySignature, secondaryPath, secondarySignature].join('\0');
}

function cleanupRetimedSubtitleCache(): void {
  for (const entry of retimedSubtitleCache.values()) {
    try {
      rmSync(entry.cleanupDir, { recursive: true, force: true });
    } catch {
      // Best-effort temp cleanup.
    }
  }
  retimedSubtitleCache.clear();
}

function registerRetimedSubtitleCleanup(): void {
  if (retimedSubtitleCleanupRegistered) return;
  retimedSubtitleCleanupRegistered = true;
  process.once('exit', cleanupRetimedSubtitleCache);
}

export function clearRetimedSecondarySubtitleCache(): void {
  cleanupRetimedSubtitleCache();
}

function splitLanguageSuffix(value: string): string[] {
  const normalizedWhole = normalizeYoutubeLangCode(value);
  const tokens = value
    .split(/[^A-Za-z0-9-]+/g)
    .map((token) => normalizeYoutubeLangCode(token))
    .filter(Boolean);
  return unique([normalizedWhole, ...tokens]);
}

function languageTokenMatches(token: string, preferredLanguage: string): boolean {
  if (token === preferredLanguage) {
    return true;
  }
  if (token.startsWith(`${preferredLanguage}-`) || preferredLanguage.startsWith(`${token}-`)) {
    return true;
  }
  return isEnglishYoutubeLang(token) && isEnglishYoutubeLang(preferredLanguage);
}

function resolveLanguageRank(suffix: string, preferredLanguages: string[]): number {
  const tokens = splitLanguageSuffix(suffix);
  for (let index = 0; index < preferredLanguages.length; index += 1) {
    const preferredLanguage = preferredLanguages[index]!;
    if (tokens.some((token) => languageTokenMatches(token, preferredLanguage))) {
      return index;
    }
  }
  return Number.POSITIVE_INFINITY;
}

function extensionRank(ext: string): number {
  if (ext === '.srt') return 0;
  if (ext === '.vtt') return 1;
  if (ext === '.ass') return 2;
  if (ext === '.ssa') return 3;
  return 4;
}

function findSidecarSubtitleCandidates(
  sourcePath: string,
  preferredLanguages: string[],
): SidecarCandidate[] {
  const source = path.parse(sourcePath);
  let entries: string[];
  try {
    entries = readdirSync(source.dir);
  } catch {
    return [];
  }

  const prefix = `${source.name}.`;
  return entries
    .map((entry) => {
      const parsed = path.parse(entry);
      const ext = parsed.ext.toLowerCase();
      if (!SUPPORTED_SUBTITLE_EXTENSIONS.has(ext) || !parsed.name.startsWith(prefix)) {
        return null;
      }
      const suffix = parsed.name.slice(prefix.length);
      const languageRank = resolveLanguageRank(suffix, preferredLanguages);
      if (!Number.isFinite(languageRank)) {
        return null;
      }
      return {
        path: path.join(source.dir, entry),
        languageRank,
        extensionRank: extensionRank(ext),
        name: entry,
      };
    })
    .filter((candidate): candidate is SidecarCandidate => candidate !== null)
    .sort((left, right) => {
      if (left.languageRank !== right.languageRank) return left.languageRank - right.languageRank;
      if (left.extensionRank !== right.extensionRank)
        return left.extensionRank - right.extensionRank;
      return left.name.localeCompare(right.name);
    });
}

function combineCueText(cues: SubtitleCue[]): string {
  return unique(cues.map((cue) => cue.text.trim()).filter(Boolean))
    .join('\n')
    .trim();
}

function overlapSeconds(cue: SubtitleCue, startSeconds: number, endSeconds: number): number {
  return (
    Math.min(cue.endTime, endSeconds + TIMING_TOLERANCE_SECONDS) -
    Math.max(cue.startTime, startSeconds - TIMING_TOLERANCE_SECONDS)
  );
}

function isSameCueTiming(left: SubtitleCue, right: SubtitleCue): boolean {
  return (
    Math.abs(left.startTime - right.startTime) <= SAME_TIMING_EPSILON_SECONDS &&
    Math.abs(left.endTime - right.endTime) <= SAME_TIMING_EPSILON_SECONDS
  );
}

function compareCueTimingMatch(
  startSeconds: number,
  endSeconds: number,
  left: { cue: SubtitleCue; overlap: number },
  right: { cue: SubtitleCue; overlap: number },
): number {
  if (left.overlap !== right.overlap) {
    return right.overlap - left.overlap;
  }

  const leftStartDistance = Math.abs(left.cue.startTime - startSeconds);
  const rightStartDistance = Math.abs(right.cue.startTime - startSeconds);
  if (leftStartDistance !== rightStartDistance) {
    return leftStartDistance - rightStartDistance;
  }

  const leftEndDistance = Math.abs(left.cue.endTime - endSeconds);
  const rightEndDistance = Math.abs(right.cue.endTime - endSeconds);
  if (leftEndDistance !== rightEndDistance) {
    return leftEndDistance - rightEndDistance;
  }

  return left.cue.startTime - right.cue.startTime;
}

function findCueTextAtTiming(cues: SubtitleCue[], startMs: number, endMs: number): string {
  const startSeconds = startMs / 1000;
  const endSeconds = endMs / 1000;
  const midpointSeconds = (startSeconds + endSeconds) / 2;

  const midpointMatches = cues
    .filter(
      (cue) =>
        cue.startTime - TIMING_TOLERANCE_SECONDS <= midpointSeconds &&
        cue.endTime + TIMING_TOLERANCE_SECONDS >= midpointSeconds,
    )
    .map((cue) => ({ cue, overlap: overlapSeconds(cue, startSeconds, endSeconds) }))
    .sort((left, right) => compareCueTimingMatch(startSeconds, endSeconds, left, right));
  const [bestMidpointMatch] = midpointMatches;
  const midpointText = bestMidpointMatch
    ? combineCueText(
        midpointMatches
          .filter((match) => isSameCueTiming(match.cue, bestMidpointMatch.cue))
          .map((match) => match.cue),
      )
    : '';
  if (midpointText) {
    return midpointText;
  }

  const [bestOverlap] = cues
    .map((cue) => ({ cue, overlap: overlapSeconds(cue, startSeconds, endSeconds) }))
    .filter((entry) => entry.overlap > 0)
    .sort((left, right) => compareCueTimingMatch(startSeconds, endSeconds, left, right));
  return bestOverlap ? bestOverlap.cue.text.trim() : '';
}

function readCueTextAtTiming(filePath: string, startMs: number, endMs: number): string {
  const content = readFileSync(filePath, 'utf8');
  const cues = parseSubtitleCues(content, filePath);
  return findCueTextAtTiming(cues, startMs, endMs);
}

async function defaultRunAlass(
  alassPath: string,
  referencePath: string,
  inputPath: string,
  outputPath: string,
): Promise<CommandResult> {
  return runCommand(alassPath, [referencePath, inputPath, outputPath], RETIMED_SUBTITLE_TIMEOUT_MS);
}

async function retimeSecondarySubtitle(input: {
  alassPath: string;
  primaryPath: string;
  secondaryPath: string;
  runAlass: RetimedSubtitleCommandRunner;
}): Promise<string> {
  const key = retimedCacheKey(input.alassPath, input.primaryPath, input.secondaryPath);
  if (!key) return '';

  const cached = retimedSubtitleCache.get(key);
  if (cached?.promise) {
    return cached.promise;
  }
  if (cached && existsSync(cached.path)) {
    return cached.path;
  }
  if (cached) {
    retimedSubtitleCache.delete(key);
    try {
      rmSync(cached.cleanupDir, { recursive: true, force: true });
    } catch {}
  }

  registerRetimedSubtitleCleanup();
  const cleanupDir = mkdtempSync(path.join(os.tmpdir(), 'subminer-retimed-secondary-'));
  const parsedSecondary = path.parse(input.secondaryPath);
  const outputPath = path.join(
    cleanupDir,
    `${parsedSecondary.name}.retimed${parsedSecondary.ext || '.srt'}`,
  );

  const entry: RetimedSubtitleCacheEntry = { path: outputPath, cleanupDir };
  entry.promise = input
    .runAlass(input.alassPath, input.primaryPath, input.secondaryPath, outputPath)
    .then((result) => {
      if (!result.ok || !existsSync(outputPath)) {
        rmSync(cleanupDir, { recursive: true, force: true });
        retimedSubtitleCache.delete(key);
        return '';
      }
      entry.promise = undefined;
      return outputPath;
    })
    .catch(() => {
      rmSync(cleanupDir, { recursive: true, force: true });
      retimedSubtitleCache.delete(key);
      return '';
    });
  retimedSubtitleCache.set(key, entry);
  return entry.promise;
}

export function resolveSecondarySubtitleTextFromSidecar(input: {
  sourcePath: string;
  startMs: number;
  endMs: number;
  languages?: readonly string[];
}): string {
  if (!input.sourcePath || !existsSync(input.sourcePath)) {
    return '';
  }
  try {
    if (!statSync(input.sourcePath).isFile()) {
      return '';
    }
  } catch {
    return '';
  }

  const preferredLanguages = expandPreferredLanguages(
    input.languages,
    DEFAULT_SECONDARY_SUBTITLE_LANGUAGES,
  );
  const candidates = findSidecarSubtitleCandidates(input.sourcePath, preferredLanguages);
  for (const candidate of candidates) {
    try {
      const text = readCueTextAtTiming(candidate.path, input.startMs, input.endMs);
      if (text) {
        return text;
      }
    } catch {
      // Try the next matching sidecar.
    }
  }

  return '';
}

export async function resolveRetimedSecondarySubtitleTextFromSidecar(
  input: RetimedSecondarySubtitleInput,
): Promise<string> {
  if (!input.sourcePath || !existsSync(input.sourcePath)) {
    return '';
  }
  try {
    if (!statSync(input.sourcePath).isFile()) {
      return '';
    }
  } catch {
    return '';
  }

  const alassPath = resolveAlassPath(input.alassPath);
  if (!alassPath) return '';

  const primaryLanguages = expandPreferredLanguages(
    input.primaryLanguages,
    DEFAULT_PRIMARY_SUBTITLE_LANGUAGES,
  );
  const secondaryLanguages = expandPreferredLanguages(
    input.languages,
    DEFAULT_SECONDARY_SUBTITLE_LANGUAGES,
  );
  const primaryCandidates = findSidecarSubtitleCandidates(input.sourcePath, primaryLanguages);
  const secondaryCandidates = findSidecarSubtitleCandidates(input.sourcePath, secondaryLanguages);
  const runAlass = input.runAlass ?? defaultRunAlass;

  for (const primary of primaryCandidates) {
    for (const secondary of secondaryCandidates) {
      if (primary.path === secondary.path) continue;
      try {
        const retimedPath = await retimeSecondarySubtitle({
          alassPath,
          primaryPath: primary.path,
          secondaryPath: secondary.path,
          runAlass,
        });
        if (!retimedPath) continue;
        const text = readCueTextAtTiming(retimedPath, input.startMs, input.endMs);
        if (text) return text;
      } catch {
        // Try the next sidecar pair.
      }
    }
  }

  return '';
}
