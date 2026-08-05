import { HAN_REGEXP_CLASS_BODY } from '../../core/text/han-code-points';
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

// Kana, halfwidth included: one of these can stand alone as a name, where a
// latin letter or a digit cannot.
const SINGLE_KANA_CHARACTER = /^[\u3040-\u30ff\u31f0-\u31ff\uff66-\uff9f]$/u;

// AniList disambiguates unnamed mob characters with a trailing letter (女子A /
// "Joshi A"), and a lone letter romanizes into a single-kana alias (A → ア)
// that collides with interjections (あ〜 matching ア). That letter is a label,
// not a name, so it is dropped where a name splits into it and before it can
// become a kana alias. A name that is genuinely one character, a character
// actually called あ or a single kanji, is a real lookup target and is kept.
function isNameDisambiguatorLetter(name: string): boolean {
  return [...name].length === 1 && !containsKanji(name) && !SINGLE_KANA_CHARACTER.test(name);
}

function isUsableNameTerm(name: string): boolean {
  return !isNameDisambiguatorLetter(name);
}

// Kana, Han (shared ranges), and the marks that only ever appear inside a
// Japanese name: iteration marks and the small ka/ke used in place names.
const JAPANESE_NAME_CHARACTERS = new RegExp(
  `^[\\u3040-\\u30ff${HAN_REGEXP_CLASS_BODY}\u3005\u3006\u30f5\u30f6\u30fc]+$`,
  'u',
);

export function isJapaneseNameSplitCandidate(name: string): boolean {
  const compact = name.replace(/[\s\u3000・･·•]/g, '');
  return containsKanji(compact) && JAPANESE_NAME_CHARACTERS.test(compact);
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
        for (const part of split) {
          if (isUsableNameTerm(part)) {
            target.add(part);
          }
        }
      }

      const splitByMiddleDot = name
        .split(/[・･·•]/)
        .map((part) => part.trim())
        .filter((part) => part.length > 0);
      if (splitByMiddleDot.length >= 2) {
        for (const part of splitByMiddleDot) {
          if (isUsableNameTerm(part)) {
            target.add(part);
          }
        }
      }

      if (target === base) {
        addJapaneseNameParts(character, name, base, resolvedSplits);
      }
    }
  }

  // Romanized forms that are a bare letter would become a single-kana alias.
  for (const alias of addRomanizedKanaAliases(
    [...romanizedBase].filter((entry) => !isNameDisambiguatorLetter(entry)),
  )) {
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
    // Only labels split off a longer name are filtered (see above); an explicit
    // one-character name reaches this point intact.
    if (isNameDisambiguatorLetter(entry)) continue;
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
