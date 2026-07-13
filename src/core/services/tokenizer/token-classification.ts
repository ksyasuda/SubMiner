import { MergedToken, PartOfSpeech } from '../../../types';

const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;

export function normalizeKana(text: string): string {
  const raw = text.trim();
  if (!raw) {
    return '';
  }

  let normalized = '';
  for (const char of raw) {
    const code = char.codePointAt(0);
    if (code === undefined) {
      continue;
    }

    normalized +=
      code >= KATAKANA_CODEPOINT_START && code <= KATAKANA_CODEPOINT_END
        ? String.fromCodePoint(code - KATAKANA_TO_HIRAGANA_OFFSET)
        : char;
  }

  return normalized;
}

export function isKanaChar(char: string): boolean {
  const code = char.codePointAt(0);
  if (code === undefined) {
    return false;
  }

  return (
    (code >= 0x3041 && code <= 0x3096) ||
    (code >= 0x309b && code <= 0x309f) ||
    code === 0x30fc ||
    (code >= 0x30a1 && code <= 0x30fa) ||
    (code >= 0x30fd && code <= 0x30ff)
  );
}

export function isKanaOnlyText(text: string | null | undefined): boolean {
  if (typeof text !== 'string') {
    return false;
  }

  const normalized = normalizeKana(text);
  return normalized.length > 0 && [...normalized].every(isKanaChar);
}

export function isKanaCandidateIgnorableChar(char: string): boolean {
  return /^[\s.,!?;:()[\]{}"'`、。！？…‥・「」『』（）［］｛｝〈〉《》【】―-]$/u.test(char);
}

export function isKanaCandidateText(text: string): boolean {
  const normalized = text.trim();
  if (normalized.length === 0) {
    return false;
  }

  let hasKana = false;
  for (const char of normalized) {
    if (isKanaChar(char)) {
      hasKana = true;
      continue;
    }
    if (!isKanaCandidateIgnorableChar(char)) {
      return false;
    }
  }

  return hasKana;
}

export function normalizePosTag(value: string | null | undefined): string {
  return typeof value === 'string' ? value.trim() : '';
}

export function splitPosTag(value: string | null | undefined): string[] {
  const normalized = normalizePosTag(value);
  if (!normalized) {
    return [];
  }

  return normalized
    .split('|')
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
}

export function isPosTagExcluded(
  value: string | null | undefined,
  exclusions: ReadonlySet<string>,
): boolean {
  const parts = splitPosTag(value);
  return parts.length > 0 && parts.every((part) => exclusions.has(part));
}

function hasKanjiChar(text: string): boolean {
  for (const char of text) {
    const code = char.codePointAt(0);
    if (
      code !== undefined &&
      ((code >= 0x3400 && code <= 0x4dbf) ||
        (code >= 0x4e00 && code <= 0x9fff) ||
        (code >= 0xf900 && code <= 0xfaff))
    ) {
      return true;
    }
  }
  return false;
}

// MeCab's 非自立 tag suppresses kana grammar nouns (こと, もの, とき), but
// Yomitan can segment kanji-bearing nouns (日, 方, 上, …) as real vocabulary.
function isKanjiNonIndependentNounToken(
  token: MergedToken,
  pos1Exclusions: ReadonlySet<string>,
): boolean {
  if (pos1Exclusions.has('名詞')) {
    return false;
  }

  const pos1Parts = splitPosTag(token.pos1);
  const pos2Parts = splitPosTag(token.pos2);
  return (
    pos1Parts.length === 1 &&
    pos1Parts[0] === '名詞' &&
    pos2Parts.length === 1 &&
    pos2Parts[0] === '非自立' &&
    (hasKanjiChar(token.surface) || hasKanjiChar(token.headword))
  );
}

export function isTokenPos2Excluded(
  token: MergedToken,
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): boolean {
  return (
    isPosTagExcluded(token.pos2, pos2Exclusions) &&
    !isKanjiNonIndependentNounToken(token, pos1Exclusions)
  );
}

export function isContentTokenByPos(
  token: MergedToken,
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): boolean {
  if (
    isPosTagExcluded(token.pos1, pos1Exclusions) ||
    isTokenPos2Excluded(token, pos1Exclusions, pos2Exclusions)
  ) {
    return false;
  }

  if (splitPosTag(token.pos1).length > 0 || splitPosTag(token.pos2).length > 0) {
    return true;
  }

  return ![PartOfSpeech.particle, PartOfSpeech.bound_auxiliary, PartOfSpeech.symbol].includes(
    token.partOfSpeech,
  );
}
