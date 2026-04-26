import * as fs from 'fs';
import * as path from 'path';
import type { AnilistMediaGuess } from '../../core/services/anilist/anilist-updater';
import { ensureDir } from '../../shared/fs-utils';

export type CharacterDictionaryManualSelection = {
  seriesKey: string;
  mediaId: number;
  mediaTitle: string;
  staleMediaIds: number[];
};

type ManualSelectionStoreFile = {
  overrides?: CharacterDictionaryManualSelection[];
};

function normalizeManualMediaId(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value)) return null;
  const mediaId = Math.floor(value);
  return mediaId > 0 ? mediaId : null;
}

function normalizeSeriesKeyPart(value: string): string {
  return value
    .normalize('NFKD')
    .replace(/[':]/g, '')
    .replace(/&/g, ' and ')
    .replace(/[^a-zA-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-{2,}/g, '-')
    .toLowerCase();
}

function dedupeNumbers(values: number[]): number[] {
  const seen = new Set<number>();
  const result: number[] = [];
  for (const value of values) {
    const normalized = normalizeManualMediaId(value);
    if (normalized === null || seen.has(normalized)) continue;
    seen.add(normalized);
    result.push(normalized);
  }
  return result;
}

function normalizeOverride(value: unknown): CharacterDictionaryManualSelection | null {
  if (!value || typeof value !== 'object') return null;
  const raw = value as Partial<CharacterDictionaryManualSelection>;
  const seriesKey = typeof raw.seriesKey === 'string' ? raw.seriesKey.trim() : '';
  const mediaId = normalizeManualMediaId(raw.mediaId);
  const mediaTitle = typeof raw.mediaTitle === 'string' ? raw.mediaTitle.trim() : '';
  if (!seriesKey || mediaId === null || !mediaTitle) return null;
  return {
    seriesKey,
    mediaId,
    mediaTitle,
    staleMediaIds: dedupeNumbers(Array.isArray(raw.staleMediaIds) ? raw.staleMediaIds : []),
  };
}

function readOverrides(filePath: string): CharacterDictionaryManualSelection[] {
  try {
    const parsed = JSON.parse(fs.readFileSync(filePath, 'utf8')) as ManualSelectionStoreFile;
    if (!Array.isArray(parsed.overrides)) return [];
    const byKey = new Map<string, CharacterDictionaryManualSelection>();
    for (const value of parsed.overrides) {
      const normalized = normalizeOverride(value);
      if (normalized) byKey.set(normalized.seriesKey, normalized);
    }
    return [...byKey.values()];
  } catch {
    return [];
  }
}

function writeOverrides(filePath: string, overrides: CharacterDictionaryManualSelection[]): void {
  ensureDir(path.dirname(filePath));
  fs.writeFileSync(filePath, JSON.stringify({ overrides }, null, 2), 'utf8');
}

export function buildCharacterDictionarySeriesKey(input: {
  mediaPath: string | null;
  mediaTitle: string | null;
  guess: AnilistMediaGuess | null;
}): string {
  const guessedTitle = input.guess?.title.trim() || input.guess?.alternativeTitle?.trim() || '';
  const sourceTitle =
    guessedTitle ||
    (input.mediaTitle && input.mediaTitle.trim()) ||
    (input.mediaPath && path.basename(input.mediaPath).replace(/\.[^.]+$/, '')) ||
    'unknown';
  const withoutEpisode = sourceTitle
    .replace(/\bS\d{1,2}E\d{1,3}\b/gi, ' ')
    .replace(/\bepisode\s+\d+\b/gi, ' ')
    .trim();
  const base = normalizeSeriesKeyPart(withoutEpisode) || 'unknown';
  return input.guess?.year ? `${base}-${input.guess.year}` : base;
}

export function createCharacterDictionaryManualSelectionStore(deps: { userDataPath: string }) {
  const filePath = path.join(deps.userDataPath, 'character-dictionaries', 'anilist-overrides.json');

  return {
    getOverride: async (seriesKey: string): Promise<CharacterDictionaryManualSelection | null> => {
      return readOverrides(filePath).find((entry) => entry.seriesKey === seriesKey) ?? null;
    },
    setOverride: async (selection: CharacterDictionaryManualSelection): Promise<void> => {
      const normalized = normalizeOverride(selection);
      if (!normalized) {
        throw new Error('Invalid character dictionary manual selection.');
      }
      const remaining = readOverrides(filePath).filter(
        (entry) => entry.seriesKey !== normalized.seriesKey,
      );
      writeOverrides(filePath, [...remaining, normalized]);
    },
  };
}
