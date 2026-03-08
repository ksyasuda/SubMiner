import fs from 'node:fs';
import path from 'node:path';

import type { SubtitleCandidate } from '../types.js';
import { YOUTUBE_SUB_EXTENSIONS } from '../types.js';
import { escapeRegExp, runExternalCommand } from '../util.js';

function filenameHasLanguageTag(filenameLower: string, langCode: string): boolean {
  const escaped = escapeRegExp(langCode);
  const pattern = new RegExp(`(^|[._-])${escaped}([._-]|$)`);
  return pattern.test(filenameLower);
}

function classifyLanguage(
  filename: string,
  primaryLangCodes: string[],
  secondaryLangCodes: string[],
): 'primary' | 'secondary' | null {
  const lower = filename.toLowerCase();
  const primary = primaryLangCodes.some((code) => filenameHasLanguageTag(lower, code));
  const secondary = secondaryLangCodes.some((code) => filenameHasLanguageTag(lower, code));
  if (primary && !secondary) return 'primary';
  if (secondary && !primary) return 'secondary';
  return null;
}

export function toYtdlpLangPattern(langCodes: string[]): string {
  return langCodes.map((lang) => `${lang}.*`).join(',');
}

export function scanSubtitleCandidates(
  tempDir: string,
  knownSet: Set<string>,
  source: SubtitleCandidate['source'],
  primaryLangCodes: string[],
  secondaryLangCodes: string[],
): SubtitleCandidate[] {
  const entries = fs.readdirSync(tempDir);
  const out: SubtitleCandidate[] = [];
  for (const name of entries) {
    const fullPath = path.join(tempDir, name);
    if (knownSet.has(fullPath)) continue;
    let stat: fs.Stats;
    try {
      stat = fs.statSync(fullPath);
    } catch {
      continue;
    }
    if (!stat.isFile()) continue;
    const ext = path.extname(fullPath).toLowerCase();
    if (!YOUTUBE_SUB_EXTENSIONS.has(ext)) continue;
    const lang = classifyLanguage(name, primaryLangCodes, secondaryLangCodes);
    if (!lang) continue;
    out.push({ path: fullPath, lang, ext, size: stat.size, source });
  }
  return out;
}

export function pickBestCandidate(candidates: SubtitleCandidate[]): SubtitleCandidate | null {
  if (candidates.length === 0) return null;
  const scored = [...candidates].sort((a, b) => {
    const srtA = a.ext === '.srt' ? 1 : 0;
    const srtB = b.ext === '.srt' ? 1 : 0;
    if (srtA !== srtB) return srtB - srtA;
    return b.size - a.size;
  });
  return scored[0] ?? null;
}

export async function downloadManualSubtitles(
  target: string,
  tempDir: string,
  langPattern: string,
  logLevel: import('../types.js').LogLevel,
  childTracker?: Set<ReturnType<typeof import('node:child_process').spawn>>,
): Promise<void> {
  await runExternalCommand(
    'yt-dlp',
    [
      '--skip-download',
      '--no-warnings',
      '--write-subs',
      '--sub-format',
      'srt/vtt/best',
      '--sub-langs',
      langPattern,
      '-o',
      path.join(tempDir, '%(id)s.%(ext)s'),
      target,
    ],
    {
      allowFailure: true,
      logLevel,
      commandLabel: 'yt-dlp:manual-subs',
      streamOutput: true,
    },
    childTracker,
  );
}
