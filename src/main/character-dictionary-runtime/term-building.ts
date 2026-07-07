import { HONORIFIC_SUFFIXES } from './constants';
import {
  addRomanizedKanaAliases,
  buildReading,
  buildReadingFromRomanized,
  containsKanji,
  hasKanaOnly,
  isRomanizedName,
  splitJapaneseName,
  splitJapaneseNameCandidates,
} from './name-reading';
import type {
  CharacterDictionaryGlossaryEntry,
  CharacterDictionaryRole,
  CharacterDictionaryTermEntry,
  CharacterRecord,
  JapaneseNameParts,
  NameReadings,
  ResolvedNameSplits,
} from './types';

export function expandRawNameVariants(rawName: string): string[] {
  const trimmed = rawName.trim();
  if (!trimmed) return [];

  const variants = new Set<string>([trimmed]);
  const outer = trimmed
    .replace(/[（(][^()（）]+[)）]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (outer && outer !== trimmed) {
    variants.add(outer);
  }

  for (const match of trimmed.matchAll(/[（(]([^()（）]+)[)）]/g)) {
    const inner = match[1]?.trim() || '';
    if (inner) {
      variants.add(inner);
    }
  }

  return [...variants];
}

export function isJapaneseNameSplitCandidate(name: string): boolean {
  const compact = name.replace(/[\s\u3000・･·•]/g, '');
  return (
    containsKanji(compact) && /^[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff々〆ヵヶー]+$/.test(compact)
  );
}

function addJapaneseNameParts(
  character: CharacterRecord,
  name: string,
  terms: Set<string>,
  resolvedSplits?: ResolvedNameSplits,
): void {
  if (!isJapaneseNameSplitCandidate(name)) return;

  const candidates = splitJapaneseNameCandidates(
    name,
    character.firstNameHint,
    character.lastNameHint,
    resolvedSplits,
  );
  for (const nameParts of candidates) {
    if (nameParts.family) {
      terms.add(nameParts.family);
    }
    if (nameParts.given) {
      terms.add(nameParts.given);
    }
  }
}

export function buildNameTerms(
  character: CharacterRecord,
  resolvedSplits?: ResolvedNameSplits,
): string[] {
  const base = new Set<string>();
  const romanizedBase = new Set<string>();
  const rawNames = [character.nativeName, character.fullName, ...character.alternativeNames];
  for (const rawName of rawNames) {
    for (const name of expandRawNameVariants(rawName)) {
      const target = isRomanizedName(name) ? romanizedBase : base;
      target.add(name);

      const compact = name.replace(/[\s\u3000]+/g, '');
      if (compact && compact !== name) {
        target.add(compact);
      }

      const noMiddleDots = compact.replace(/[・･·•]/g, '');
      if (noMiddleDots && noMiddleDots !== compact) {
        target.add(noMiddleDots);
      }

      const split = name.split(/[\s\u3000]+/).filter((part) => part.trim().length > 0);
      if (split.length === 2) {
        target.add(split[0]!);
        target.add(split[1]!);
      }

      const splitByMiddleDot = name
        .split(/[・･·•]/)
        .map((part) => part.trim())
        .filter((part) => part.length > 0);
      if (splitByMiddleDot.length >= 2) {
        for (const part of splitByMiddleDot) {
          target.add(part);
        }
      }

      if (target === base) {
        addJapaneseNameParts(character, name, base, resolvedSplits);
      }
    }
  }

  for (const alias of addRomanizedKanaAliases(romanizedBase)) {
    base.add(alias);
  }

  const nativeParts = splitJapaneseName(
    character.nativeName,
    character.firstNameHint,
    character.lastNameHint,
    resolvedSplits,
  );
  if (nativeParts.family) {
    base.add(nativeParts.family);
  }
  if (nativeParts.given) {
    base.add(nativeParts.given);
  }

  const withHonorifics = new Set<string>();
  for (const entry of base) {
    withHonorifics.add(entry);
    for (const suffix of HONORIFIC_SUFFIXES) {
      withHonorifics.add(`${entry}${suffix.term}`);
    }
  }

  return [...withHonorifics].filter((entry) => entry.trim().length > 0);
}

export function buildVisibleNameTerms(nameTerms: string[]): string[] {
  const allTerms = new Set(nameTerms);
  return nameTerms.filter((term) => {
    for (const suffix of HONORIFIC_SUFFIXES) {
      if (!term.endsWith(suffix.term) || term.length <= suffix.term.length) {
        continue;
      }
      if (allTerms.has(term.slice(0, -suffix.term.length))) {
        return false;
      }
    }
    return true;
  });
}

export function buildReadingForTerm(
  term: string,
  character: CharacterRecord,
  readings: NameReadings,
  nameParts: JapaneseNameParts,
): string {
  for (const suffix of HONORIFIC_SUFFIXES) {
    if (term.endsWith(suffix.term) && term.length > suffix.term.length) {
      const baseTerm = term.slice(0, -suffix.term.length);
      const baseReading = buildReadingForTerm(baseTerm, character, readings, nameParts);
      return baseReading ? `${baseReading}${suffix.reading}` : '';
    }
  }

  const compactNative = character.nativeName.replace(/[\s\u3000]+/g, '');
  const noMiddleDotsNative = compactNative.replace(/[・･·•]/g, '');
  if (
    term === character.nativeName ||
    term === compactNative ||
    term === noMiddleDotsNative ||
    term === nameParts.original ||
    term === nameParts.combined
  ) {
    return readings.full;
  }

  const familyCompact = nameParts.family?.replace(/[・･·•]/g, '') || '';
  if (nameParts.family && (term === nameParts.family || term === familyCompact)) {
    return readings.family;
  }

  const givenCompact = nameParts.given?.replace(/[・･·•]/g, '') || '';
  if (nameParts.given && (term === nameParts.given || term === givenCompact)) {
    return readings.given;
  }

  const compact = term.replace(/[\s\u3000]+/g, '');
  if (hasKanaOnly(compact)) {
    return buildReading(compact);
  }

  if (isRomanizedName(term)) {
    return buildReadingFromRomanized(term) || readings.full;
  }

  return '';
}

function roleInfo(role: CharacterDictionaryRole): { tag: string; score: number } {
  if (role === 'main') return { tag: 'main', score: 100 };
  if (role === 'primary') return { tag: 'primary', score: 75 };
  if (role === 'side') return { tag: 'side', score: 50 };
  return { tag: 'appears', score: 25 };
}

export function buildTermEntry(
  term: string,
  reading: string,
  role: CharacterDictionaryRole,
  glossary: CharacterDictionaryGlossaryEntry[],
): CharacterDictionaryTermEntry {
  const { tag, score } = roleInfo(role);
  return [term, reading, `name ${tag}`, '', score, glossary, 0, ''];
}
