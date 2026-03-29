import { HONORIFIC_SUFFIXES } from './constants';
import {
  addRomanizedKanaAliases,
  buildReading,
  buildReadingFromRomanized,
  hasKanaOnly,
  isRomanizedName,
  splitJapaneseName,
} from './name-reading';
import type {
  CharacterDictionaryGlossaryEntry,
  CharacterDictionaryRole,
  CharacterDictionaryTermEntry,
  CharacterRecord,
  JapaneseNameParts,
  NameReadings,
} from './types';

function expandRawNameVariants(rawName: string): string[] {
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

export function buildNameTerms(character: CharacterRecord): string[] {
  const base = new Set<string>();
  const rawNames = [character.nativeName, character.fullName, ...character.alternativeNames];
  for (const rawName of rawNames) {
    for (const name of expandRawNameVariants(rawName)) {
      base.add(name);

      const compact = name.replace(/[\s\u3000]+/g, '');
      if (compact && compact !== name) {
        base.add(compact);
      }

      const noMiddleDots = compact.replace(/[・･·•]/g, '');
      if (noMiddleDots && noMiddleDots !== compact) {
        base.add(noMiddleDots);
      }

      const split = name.split(/[\s\u3000]+/).filter((part) => part.trim().length > 0);
      if (split.length === 2) {
        base.add(split[0]!);
        base.add(split[1]!);
      }

      const splitByMiddleDot = name
        .split(/[・･·•]/)
        .map((part) => part.trim())
        .filter((part) => part.length > 0);
      if (splitByMiddleDot.length >= 2) {
        for (const part of splitByMiddleDot) {
          base.add(part);
        }
      }
    }
  }

  const nativeParts = splitJapaneseName(
    character.nativeName,
    character.firstNameHint,
    character.lastNameHint,
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

  for (const alias of addRomanizedKanaAliases(withHonorifics)) {
    withHonorifics.add(alias);
    for (const suffix of HONORIFIC_SUFFIXES) {
      withHonorifics.add(`${alias}${suffix.term}`);
    }
  }

  return [...withHonorifics].filter((entry) => entry.trim().length > 0);
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
