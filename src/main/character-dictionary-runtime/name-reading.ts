import { HONORIFIC_SUFFIXES } from './constants';
import type { JapaneseNameParts, NameReadings, ResolvedNameSplits } from './types';

export function hasKanaOnly(value: string): boolean {
  return /^[\u3040-\u309f\u30a0-\u30ffー]+$/.test(value);
}

function katakanaToHiragana(value: string): string {
  let output = '';
  for (const char of value) {
    const code = char.charCodeAt(0);
    if (code >= 0x30a1 && code <= 0x30f6) {
      output += String.fromCharCode(code - 0x60);
      continue;
    }
    output += char;
  }
  return output;
}

export function buildReading(term: string): string {
  const compact = term.replace(/\s+/g, '').trim();
  if (!compact || !hasKanaOnly(compact)) {
    return '';
  }
  return katakanaToHiragana(compact);
}

export function containsKanji(value: string): boolean {
  for (const char of value) {
    const code = char.charCodeAt(0);
    if ((code >= 0x4e00 && code <= 0x9fff) || (code >= 0x3400 && code <= 0x4dbf)) {
      return true;
    }
  }
  return false;
}

export function isRomanizedName(value: string): boolean {
  return /^[A-Za-zĀĪŪĒŌÂÊÎÔÛāīūēōâêîôû'’.\-\s]+$/.test(value);
}

function normalizeRomanizedName(value: string): string {
  return value
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[’']/g, '')
    .replace(/[.\-]/g, ' ')
    .replace(/ā|â/g, 'aa')
    .replace(/ī|î/g, 'ii')
    .replace(/ū|û/g, 'uu')
    .replace(/ē|ê/g, 'ei')
    .replace(/ō|ô/g, 'ou')
    .replace(/\s+/g, ' ')
    .trim();
}

const ROMANIZED_KANA_DIGRAPHS: ReadonlyArray<[string, string]> = [
  ['kya', 'キャ'],
  ['kyu', 'キュ'],
  ['kyo', 'キョ'],
  ['gya', 'ギャ'],
  ['gyu', 'ギュ'],
  ['gyo', 'ギョ'],
  ['sha', 'シャ'],
  ['shu', 'シュ'],
  ['sho', 'ショ'],
  ['sya', 'シャ'],
  ['syu', 'シュ'],
  ['syo', 'ショ'],
  ['ja', 'ジャ'],
  ['ju', 'ジュ'],
  ['jo', 'ジョ'],
  ['jya', 'ジャ'],
  ['jyu', 'ジュ'],
  ['jyo', 'ジョ'],
  ['cha', 'チャ'],
  ['chu', 'チュ'],
  ['cho', 'チョ'],
  ['tya', 'チャ'],
  ['tyu', 'チュ'],
  ['tyo', 'チョ'],
  ['cya', 'チャ'],
  ['cyu', 'チュ'],
  ['cyo', 'チョ'],
  ['nya', 'ニャ'],
  ['nyu', 'ニュ'],
  ['nyo', 'ニョ'],
  ['hya', 'ヒャ'],
  ['hyu', 'ヒュ'],
  ['hyo', 'ヒョ'],
  ['bya', 'ビャ'],
  ['byu', 'ビュ'],
  ['byo', 'ビョ'],
  ['pya', 'ピャ'],
  ['pyu', 'ピュ'],
  ['pyo', 'ピョ'],
  ['mya', 'ミャ'],
  ['myu', 'ミュ'],
  ['myo', 'ミョ'],
  ['rya', 'リャ'],
  ['ryu', 'リュ'],
  ['ryo', 'リョ'],
  ['fa', 'ファ'],
  ['fi', 'フィ'],
  ['fe', 'フェ'],
  ['fo', 'フォ'],
  ['fyu', 'フュ'],
  ['fyo', 'フョ'],
  ['fya', 'フャ'],
  ['va', 'ヴァ'],
  ['vi', 'ヴィ'],
  ['vu', 'ヴ'],
  ['ve', 'ヴェ'],
  ['vo', 'ヴォ'],
  ['she', 'シェ'],
  ['che', 'チェ'],
  ['je', 'ジェ'],
  ['tsi', 'ツィ'],
  ['tse', 'ツェ'],
  ['tsa', 'ツァ'],
  ['tso', 'ツォ'],
  ['thi', 'ティ'],
  ['thu', 'テュ'],
  ['dhi', 'ディ'],
  ['dhu', 'デュ'],
  ['wi', 'ウィ'],
  ['we', 'ウェ'],
  ['wo', 'ウォ'],
];

const ROMANIZED_KANA_MONOGRAPHS: ReadonlyArray<[string, string]> = [
  ['a', 'ア'],
  ['i', 'イ'],
  ['u', 'ウ'],
  ['e', 'エ'],
  ['o', 'オ'],
  ['ka', 'カ'],
  ['ki', 'キ'],
  ['ku', 'ク'],
  ['ke', 'ケ'],
  ['ko', 'コ'],
  ['ga', 'ガ'],
  ['gi', 'ギ'],
  ['gu', 'グ'],
  ['ge', 'ゲ'],
  ['go', 'ゴ'],
  ['sa', 'サ'],
  ['shi', 'シ'],
  ['si', 'シ'],
  ['su', 'ス'],
  ['se', 'セ'],
  ['so', 'ソ'],
  ['za', 'ザ'],
  ['ji', 'ジ'],
  ['zi', 'ジ'],
  ['zu', 'ズ'],
  ['ze', 'ゼ'],
  ['zo', 'ゾ'],
  ['ta', 'タ'],
  ['chi', 'チ'],
  ['ti', 'チ'],
  ['tsu', 'ツ'],
  ['tu', 'ツ'],
  ['te', 'テ'],
  ['to', 'ト'],
  ['da', 'ダ'],
  ['de', 'デ'],
  ['do', 'ド'],
  ['na', 'ナ'],
  ['ni', 'ニ'],
  ['nu', 'ヌ'],
  ['ne', 'ネ'],
  ['no', 'ノ'],
  ['ha', 'ハ'],
  ['hi', 'ヒ'],
  ['fu', 'フ'],
  ['hu', 'フ'],
  ['he', 'ヘ'],
  ['ho', 'ホ'],
  ['ba', 'バ'],
  ['bi', 'ビ'],
  ['bu', 'ブ'],
  ['be', 'ベ'],
  ['bo', 'ボ'],
  ['pa', 'パ'],
  ['pi', 'ピ'],
  ['pu', 'プ'],
  ['pe', 'ペ'],
  ['po', 'ポ'],
  ['ma', 'マ'],
  ['mi', 'ミ'],
  ['mu', 'ム'],
  ['me', 'メ'],
  ['mo', 'モ'],
  ['ya', 'ヤ'],
  ['yu', 'ユ'],
  ['yo', 'ヨ'],
  ['ra', 'ラ'],
  ['ri', 'リ'],
  ['ru', 'ル'],
  ['re', 'レ'],
  ['ro', 'ロ'],
  ['wa', 'ワ'],
  ['w', 'ウ'],
  ['wo', 'ヲ'],
  ['n', 'ン'],
];

function romanizedTokenToKatakana(token: string): string | null {
  const normalized = normalizeRomanizedName(token).replace(/\s+/g, '');
  if (!normalized || !/^[a-z]+$/.test(normalized)) {
    return null;
  }

  let output = '';
  for (let i = 0; i < normalized.length; ) {
    const current = normalized[i]!;
    const next = normalized[i + 1] ?? '';

    if (
      i + 1 < normalized.length &&
      current === next &&
      current !== 'n' &&
      !'aeiou'.includes(current)
    ) {
      output += 'ッ';
      i += 1;
      continue;
    }

    if (current === 'n' && next.length > 0 && next !== 'y' && !'aeiou'.includes(next)) {
      output += 'ン';
      i += 1;
      continue;
    }

    const digraph = ROMANIZED_KANA_DIGRAPHS.find(([romaji]) => normalized.startsWith(romaji, i));
    if (digraph) {
      output += digraph[1];
      i += digraph[0].length;
      continue;
    }

    const monograph = ROMANIZED_KANA_MONOGRAPHS.find(([romaji]) =>
      normalized.startsWith(romaji, i),
    );
    if (monograph) {
      output += monograph[1];
      i += monograph[0].length;
      continue;
    }

    return null;
  }

  return output.length > 0 ? output : null;
}

export function buildReadingFromRomanized(value: string): string {
  const katakana = romanizedTokenToKatakana(value);
  return katakana ? katakanaToHiragana(katakana) : '';
}

export function buildReadingFromHint(value: string): string {
  return buildReading(value) || buildReadingFromRomanized(value);
}

function scoreJapaneseNamePartLength(length: number): number {
  if (length === 2) return 3;
  if (length === 1 || length === 3) return 2;
  if (length === 4) return 1;
  return 0;
}

// Ranks every possible family/given boundary. Reading-length ratios cannot
// always identify the true boundary (あずま can be one kanji or two), so
// callers may take the top candidates rather than trusting only the best.
function inferJapaneseNameSplitIndices(
  nameOriginal: string,
  firstNameHint: string,
  lastNameHint: string,
): number[] {
  const chars = [...nameOriginal];
  if (chars.length < 2) return [];

  const familyHintLength = [...buildReadingFromHint(lastNameHint)].length;
  const givenHintLength = [...buildReadingFromHint(firstNameHint)].length;
  const totalHintLength = familyHintLength + givenHintLength;
  const defaultBoundary = Math.round(chars.length / 2);
  const scored: Array<{ index: number; score: number }> = [];

  for (let index = 1; index < chars.length; index += 1) {
    const familyLength = index;
    const givenLength = chars.length - index;
    let score =
      scoreJapaneseNamePartLength(familyLength) + scoreJapaneseNamePartLength(givenLength);

    if (chars.length >= 4 && familyLength >= 2 && givenLength >= 2) {
      score += 1;
    }

    if (totalHintLength > 0) {
      const expectedFamilyLength = (chars.length * familyHintLength) / totalHintLength;
      score -= Math.abs(familyLength - expectedFamilyLength) * 1.5;
    } else {
      score -= Math.abs(familyLength - defaultBoundary) * 0.5;
    }

    if (familyLength === givenLength) {
      score += 0.25;
    }

    scored.push({ index, score });
  }

  return scored.sort((left, right) => right.score - left.score).map((entry) => entry.index);
}

export function addRomanizedKanaAliases(values: Iterable<string>): string[] {
  const aliases = new Set<string>();
  for (const value of values) {
    const trimmed = value.trim();
    if (!trimmed || !isRomanizedName(trimmed)) continue;
    const katakana = romanizedTokenToKatakana(trimmed);
    if (katakana) {
      aliases.add(katakana);
    }
  }
  return [...aliases];
}

export function splitJapaneseName(
  nameOriginal: string,
  firstNameHint?: string,
  lastNameHint?: string,
  resolvedSplits?: ResolvedNameSplits,
): JapaneseNameParts {
  const trimmed = nameOriginal.trim();
  if (!trimmed) {
    return {
      hasSpace: false,
      original: '',
      combined: '',
      family: null,
      given: null,
    };
  }

  const normalizedSpace = trimmed.replace(/[\s\u3000]+/g, ' ').trim();
  const spaceParts = normalizedSpace.split(' ').filter((part) => part.length > 0);
  if (spaceParts.length === 2) {
    const family = spaceParts[0]!;
    const given = spaceParts[1]!;
    return {
      hasSpace: true,
      original: normalizedSpace,
      combined: `${family}${given}`,
      family,
      given,
    };
  }

  const middleDotParts = trimmed
    .split(/[・･·•]/)
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
  if (middleDotParts.length === 2) {
    const family = middleDotParts[0]!;
    const given = middleDotParts[1]!;
    return {
      hasSpace: true,
      original: trimmed,
      combined: `${family}${given}`,
      family,
      given,
    };
  }

  const resolvedSplit = resolvedSplits?.get(trimmed);
  if (
    resolvedSplit &&
    resolvedSplit.family &&
    resolvedSplit.given &&
    `${resolvedSplit.family}${resolvedSplit.given}` === trimmed
  ) {
    return {
      hasSpace: true,
      original: trimmed,
      combined: trimmed,
      family: resolvedSplit.family,
      given: resolvedSplit.given,
    };
  }

  const hintedFirst = firstNameHint?.trim() || '';
  const hintedLast = lastNameHint?.trim() || '';
  if (hintedFirst && hintedLast) {
    const familyGiven = `${hintedLast}${hintedFirst}`;
    if (trimmed === familyGiven) {
      return {
        hasSpace: true,
        original: trimmed,
        combined: familyGiven,
        family: hintedLast,
        given: hintedFirst,
      };
    }

    const givenFamily = `${hintedFirst}${hintedLast}`;
    if (trimmed === givenFamily) {
      return {
        hasSpace: true,
        original: trimmed,
        combined: givenFamily,
        family: hintedFirst,
        given: hintedLast,
      };
    }
  }

  if (hintedFirst && hintedLast && containsKanji(trimmed)) {
    const splitIndex = inferJapaneseNameSplitIndices(trimmed, hintedFirst, hintedLast)[0] ?? null;
    if (splitIndex != null) {
      const chars = [...trimmed];
      const family = chars.slice(0, splitIndex).join('');
      const given = chars.slice(splitIndex).join('');
      if (family && given) {
        return {
          hasSpace: true,
          original: trimmed,
          combined: trimmed,
          family,
          given,
        };
      }
    }
  }

  return {
    hasSpace: false,
    original: trimmed,
    combined: trimmed,
    family: null,
    given: null,
  };
}

const MAX_INFERRED_SPLIT_CANDIDATES = 2;

// Returns the possible family/given splits, best first. Only a boundary
// guessed by the length heuristic is ambiguous (あずま can be one kanji or
// two), so only that path yields a runner-up candidate; explicit separators,
// resolved (MeCab) splits, and exact hint matches are trusted as-is.
export function splitJapaneseNameCandidates(
  nameOriginal: string,
  firstNameHint?: string,
  lastNameHint?: string,
  resolvedSplits?: ResolvedNameSplits,
): JapaneseNameParts[] {
  const primary = splitJapaneseName(nameOriginal, firstNameHint, lastNameHint, resolvedSplits);
  if (!primary.family || !primary.given) {
    return [primary];
  }

  const trimmed = nameOriginal.trim();
  if (primary.combined !== trimmed) {
    return [primary];
  }
  const resolvedSplit = resolvedSplits?.get(trimmed);
  if (resolvedSplit && `${resolvedSplit.family}${resolvedSplit.given}` === trimmed) {
    return [primary];
  }
  const hintedFirst = firstNameHint?.trim() || '';
  const hintedLast = lastNameHint?.trim() || '';
  if (`${hintedLast}${hintedFirst}` === trimmed || `${hintedFirst}${hintedLast}` === trimmed) {
    return [primary];
  }

  const candidates = [primary];
  const chars = [...trimmed];
  const splitIndices = inferJapaneseNameSplitIndices(trimmed, hintedFirst, hintedLast);
  for (const splitIndex of splitIndices.slice(1, MAX_INFERRED_SPLIT_CANDIDATES)) {
    const family = chars.slice(0, splitIndex).join('');
    const given = chars.slice(splitIndex).join('');
    if (family && given) {
      candidates.push({
        hasSpace: true,
        original: trimmed,
        combined: trimmed,
        family,
        given,
      });
    }
  }
  return candidates;
}

export function generateNameReadings(
  nameOriginal: string,
  romanizedName: string,
  firstNameHint?: string,
  lastNameHint?: string,
  resolvedSplits?: ResolvedNameSplits,
): NameReadings {
  const trimmed = nameOriginal.trim();
  if (!trimmed) {
    return {
      hasSpace: false,
      original: '',
      full: '',
      family: '',
      given: '',
    };
  }

  const nameParts = splitJapaneseName(trimmed, firstNameHint, lastNameHint, resolvedSplits);
  if (!nameParts.hasSpace || !nameParts.family || !nameParts.given) {
    const full = containsKanji(trimmed)
      ? buildReadingFromRomanized(romanizedName)
      : buildReading(trimmed);
    return {
      hasSpace: false,
      original: trimmed,
      full,
      family: full,
      given: full,
    };
  }

  const romanizedParts = romanizedName
    .trim()
    .split(/\s+/)
    .filter((part) => part.length > 0);
  const familyFromHints = buildReadingFromHint(lastNameHint || '');
  const givenFromHints = buildReadingFromHint(firstNameHint || '');
  const familyRomajiFallback = romanizedParts[0] || '';
  const givenRomajiFallback = romanizedParts.slice(1).join(' ');
  const family =
    familyFromHints ||
    (containsKanji(nameParts.family)
      ? buildReadingFromRomanized(familyRomajiFallback)
      : buildReading(nameParts.family));
  const given =
    givenFromHints ||
    (containsKanji(nameParts.given)
      ? buildReadingFromRomanized(givenRomajiFallback)
      : buildReading(nameParts.given));
  const full =
    `${family}${given}` || buildReading(trimmed) || buildReadingFromRomanized(romanizedName);

  return {
    hasSpace: true,
    original: nameParts.original,
    full,
    family,
    given,
  };
}

export function buildHonorificAliases(value: string): string[] {
  return HONORIFIC_SUFFIXES.map((suffix) => `${value}${suffix.term}`);
}
