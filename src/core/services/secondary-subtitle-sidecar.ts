import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import { parseSubtitleCues, type SubtitleCue } from './subtitle-cue-parser.js';
import { isEnglishYoutubeLang, normalizeYoutubeLangCode } from './youtube/labels.js';

const DEFAULT_SECONDARY_SUBTITLE_LANGUAGES = ['en', 'eng', 'english', 'en-us', 'enus'];
const SUPPORTED_SUBTITLE_EXTENSIONS = new Set(['.srt', '.vtt', '.ass', '.ssa']);
const TIMING_TOLERANCE_SECONDS = 0.25;
const SAME_TIMING_EPSILON_SECONDS = 0.001;

type SidecarCandidate = {
  path: string;
  languageRank: number;
  extensionRank: number;
  name: string;
};

function unique(values: string[]): string[] {
  return values.filter((value, index) => value.length > 0 && values.indexOf(value) === index);
}

function expandPreferredLanguages(languages: readonly string[] | undefined): string[] {
  const normalized = unique(
    (languages ?? []).map((language) => normalizeYoutubeLangCode(language)).filter(Boolean),
  );
  const base = normalized.length > 0 ? normalized : DEFAULT_SECONDARY_SUBTITLE_LANGUAGES;
  const expanded: string[] = [];
  for (const language of base) {
    expanded.push(language);
    if (isEnglishYoutubeLang(language)) {
      expanded.push(...DEFAULT_SECONDARY_SUBTITLE_LANGUAGES);
    }
  }
  return unique(expanded);
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

  const preferredLanguages = expandPreferredLanguages(input.languages);
  const candidates = findSidecarSubtitleCandidates(input.sourcePath, preferredLanguages);
  for (const candidate of candidates) {
    try {
      const content = readFileSync(candidate.path, 'utf8');
      const cues = parseSubtitleCues(content, candidate.path);
      const text = findCueTextAtTiming(cues, input.startMs, input.endMs);
      if (text) {
        return text;
      }
    } catch {
      // Try the next matching sidecar.
    }
  }

  return '';
}
