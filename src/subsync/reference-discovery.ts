import * as fs from 'node:fs';
import * as path from 'node:path';
import type { ReferenceAudioStatus } from './alass-align';
import { isSignsOrSongsTrackMeta } from './subtitle-lines';
import type { MpvTrack } from './utils';

const SIDECAR_SUBTITLE_EXTENSIONS = new Set(['.srt', '.ass', '.ssa', '.vtt']);

export interface SidecarSubtitleFile {
  path: string;
  filename: string;
  languageHint: string;
  isSignsOnly: boolean;
}

export interface EvaluatedSubtitleReference {
  key: string;
  kind: 'track' | 'sidecar';
  trackId: number | null;
  filePath: string | null;
  label: string;
  isSecondary: boolean;
  isEmbedded: boolean;
  isEnglish: boolean;
  isFullDialogueHint: boolean;
  isSignsOnly: boolean;
  dialogueCount: number;
  signCount: number;
  audioStatus: ReferenceAudioStatus;
  audioOffsetMs: number | null;
  audioConfidence: number;
}

function normalizePathKey(filePath: string): string {
  const resolved = path.resolve(filePath).replace(/\\/g, '/');
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved;
}

export function isEnglishReferenceMeta(meta: {
  lang?: string;
  title?: string;
  filename?: string;
}): boolean {
  const lang = (meta.lang ?? '').trim();
  if (/^(?:en|eng|english)(?:[-_]|$)/i.test(lang)) return true;
  const combined = `${meta.title ?? ''} ${meta.filename ?? ''}`;
  return /(?:^|[\s.\[(_-])(?:en|eng|english)(?=$|[\s.\])_-])/i.test(combined);
}

export function isFullDialogueHintMeta(meta: { title?: string; filename?: string }): boolean {
  const combined = `${meta.title ?? ''} ${meta.filename ?? ''}`;
  return /\b(?:full|dialogue|dialog|complete|main)\b/i.test(combined);
}

function inferLanguageHintFromFilename(filename: string, videoStem: string): string {
  const parsed = path.parse(filename);
  const remainder = parsed.name.startsWith(videoStem)
    ? parsed.name.slice(videoStem.length).replace(/^[.\s_-]+/, '')
    : parsed.name;
  return remainder || parsed.ext.slice(1);
}

/**
 * Scan the video's directory for unloaded sidecar subtitle files matching the
 * video's stem, excluding the target subtitle, `_retimed` files, and paths
 * already present in mpv's track list.
 */
export function discoverSidecarReferenceFiles(
  videoPath: string,
  loadedPaths: readonly string[],
): SidecarSubtitleFile[] {
  const parsedVideo = path.parse(videoPath);
  if (!parsedVideo.dir || !parsedVideo.name) return [];

  let entries: string[];
  try {
    entries = fs.readdirSync(parsedVideo.dir);
  } catch {
    return [];
  }

  const excluded = new Set(loadedPaths.map(normalizePathKey));
  const stemLower = parsedVideo.name.toLowerCase();
  const candidates: SidecarSubtitleFile[] = [];

  for (const entry of entries) {
    const parsedEntry = path.parse(entry);
    const ext = parsedEntry.ext.toLowerCase();
    if (!SIDECAR_SUBTITLE_EXTENSIONS.has(ext)) continue;
    const nameLower = parsedEntry.name.toLowerCase();
    if (nameLower.endsWith('_retimed')) continue;
    if (
      nameLower !== stemLower &&
      !nameLower.startsWith(`${stemLower}.`) &&
      !nameLower.startsWith(`${stemLower}_`) &&
      !nameLower.startsWith(`${stemLower}-`) &&
      !nameLower.startsWith(`${stemLower} `)
    ) {
      continue;
    }

    const fullPath = path.join(parsedVideo.dir, entry);
    if (excluded.has(normalizePathKey(fullPath))) continue;

    try {
      if (!fs.statSync(fullPath).isFile()) continue;
    } catch {
      continue;
    }

    candidates.push({
      path: fullPath,
      filename: entry,
      languageHint: inferLanguageHintFromFilename(entry, parsedVideo.name),
      isSignsOnly: isSignsOrSongsTrackMeta({ 'external-filename': fullPath }),
    });
  }

  return candidates.sort((a, b) => a.filename.localeCompare(b.filename));
}

export function formatReferenceCandidateDetail(candidate: EvaluatedSubtitleReference): string {
  const parts: string[] = [];
  if (candidate.audioStatus === 'aligned') {
    const pct = Math.round(candidate.audioConfidence * 100);
    parts.push(`aligned with audio${pct > 0 ? ` (${pct}%)` : ''}`);
  } else if (candidate.audioStatus === 'misaligned' && candidate.audioOffsetMs !== null) {
    const sec = (candidate.audioOffsetMs / 1000).toFixed(1);
    const signed = candidate.audioOffsetMs >= 0 ? `+${sec}s` : `${sec}s`;
    parts.push(`${signed} off audio`);
  }
  if (candidate.isSignsOnly) {
    parts.push('signs/songs');
  } else if (candidate.signCount > 0 && candidate.dialogueCount > 0) {
    parts.push(`${candidate.dialogueCount} dialogue, ${candidate.signCount} signs/songs stripped`);
  } else if (candidate.dialogueCount > 0) {
    parts.push(`${candidate.dialogueCount} lines`);
  }
  if (candidate.kind === 'sidecar') {
    parts.push('unloaded sidecar');
  } else if (!candidate.isSecondary && !candidate.isEmbedded) {
    parts.push('external');
  }
  return parts.join(' · ');
}

/**
 * Score candidate subtitle references so the best audio-verified dialogue track
 * wins even when it isn't the currently loaded secondary subtitle.
 */
export function scoreReferenceCandidate(candidate: EvaluatedSubtitleReference): number {
  let score = 0;

  // Audio verification dominates when available: an aligned track ALWAYS beats a misaligned one.
  if (candidate.audioStatus === 'aligned') {
    score += 10_000;
    score += Math.round(candidate.audioConfidence * 1000);
    if (candidate.audioOffsetMs !== null) {
      // Closer to 0ms offset against audio is better.
      score -= Math.min(500, Math.abs(candidate.audioOffsetMs));
    }
  } else if (candidate.audioStatus === 'unverified') {
    score += 2_000;
  } else {
    // Explicitly misaligned with audio: push well below aligned and unverified.
    score -= 5_000;
  }

  // Signs/songs-only tracks or tiny dialogue counts are poor references.
  if (candidate.isSignsOnly) score -= 4_000;
  if (candidate.dialogueCount < 20) score -= 3_000;

  // Embedded tracks in the container are usually authored for this exact video rip.
  if (candidate.isEmbedded) score += 250;
  if (candidate.isFullDialogueHint) score += 150;
  if (candidate.isEnglish) score += 100;
  if (candidate.isSecondary) score += 40;

  // Prefer dialogue-rich tracks (capped so huge karaoke tracks don't skew score).
  score += Math.min(200, Math.floor(candidate.dialogueCount / 5));

  return score;
}

export function rankReferenceCandidates(
  candidates: readonly EvaluatedSubtitleReference[],
): EvaluatedSubtitleReference[] {
  return [...candidates].sort((a, b) => scoreReferenceCandidate(b) - scoreReferenceCandidate(a));
}

export function buildTrackReferenceLabel(track: MpvTrack, isSecondary: boolean): string {
  const id = typeof track.id === 'number' ? track.id : -1;
  const source = track.external ? 'External' : 'Internal';
  const parts: string[] = [];
  if (track.lang) parts.push(track.lang);
  if (track.title && track.title !== track.lang) parts.push(track.title);
  if (parts.length === 0 && track.external && track['external-filename']) {
    parts.push(path.basename(track['external-filename']));
  }
  const descriptor = parts.length > 0 ? parts.join(' - ') : 'unknown';
  const role = isSecondary ? ' (secondary)' : '';
  return `${source} #${id} - ${descriptor}${role}`;
}
