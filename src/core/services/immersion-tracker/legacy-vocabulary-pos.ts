import type { Token } from '../../../types';
import type { LegacyVocabularyPosResolution } from './types';
import { deriveStoredPartOfSpeech } from '../tokenizer/part-of-speech';

const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;

function normalizeLookupText(value: string | null | undefined): string {
  return typeof value === 'string' ? value.trim() : '';
}

function katakanaToHiragana(text: string): string {
  let normalized = '';
  for (const char of text) {
    const code = char.codePointAt(0);
    if (code === undefined) {
      continue;
    }
    if (code >= KATAKANA_CODEPOINT_START && code <= KATAKANA_CODEPOINT_END) {
      normalized += String.fromCodePoint(code - KATAKANA_TO_HIRAGANA_OFFSET);
      continue;
    }
    normalized += char;
  }
  return normalized;
}

function toResolution(token: Token): LegacyVocabularyPosResolution {
  return {
    headword: normalizeLookupText(token.headword) || normalizeLookupText(token.word),
    reading: katakanaToHiragana(normalizeLookupText(token.katakanaReading)),
    partOfSpeech: deriveStoredPartOfSpeech({
      partOfSpeech: token.partOfSpeech,
      pos1: token.pos1,
    }),
    pos1: normalizeLookupText(token.pos1),
    pos2: normalizeLookupText(token.pos2),
    pos3: normalizeLookupText(token.pos3),
  };
}

export function resolveLegacyVocabularyPosFromTokens(
  lookupText: string,
  tokens: Token[] | null,
): LegacyVocabularyPosResolution | null {
  const normalizedLookup = normalizeLookupText(lookupText);
  if (!normalizedLookup || !tokens || tokens.length === 0) {
    return null;
  }

  const exactSurfaceMatches = tokens.filter(
    (token) => normalizeLookupText(token.word) === normalizedLookup,
  );
  if (exactSurfaceMatches.length === 1) {
    return toResolution(exactSurfaceMatches[0]!);
  }

  const exactHeadwordMatches = tokens.filter(
    (token) => normalizeLookupText(token.headword) === normalizedLookup,
  );
  if (exactHeadwordMatches.length === 1) {
    return toResolution(exactHeadwordMatches[0]!);
  }

  if (tokens.length === 1) {
    return toResolution(tokens[0]!);
  }

  return null;
}
