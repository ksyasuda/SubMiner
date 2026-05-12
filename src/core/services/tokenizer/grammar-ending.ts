const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;

const SENTENCE_FINAL_PARTICLE_SUFFIXES = ['', 'か', 'ね', 'よ', 'な', 'わ'] as const;
const EXPLANATORY_ENDING_PREFIXES = ['ん', 'の', 'なん', 'なの'] as const;
const EXPLANATORY_ENDING_CORES = [
  'だ',
  'です',
  'でした',
  'だった',
  'では',
  'じゃ',
  'でしょう',
  'だろう',
] as const;
const EXPLANATORY_ENDING_TRAILING_PARTICLES = [
  '',
  'か',
  'ね',
  'よ',
  'な',
  'けど',
  'よね',
  'かな',
  'かね',
] as const;
const EXPLANATORY_ENDING_THOUGHT_SUFFIXES = ['か', 'かな', 'かね'] as const;
const NEGATIVE_COPULA_PREFIXES = ['じゃ', 'では'] as const;

export function normalizeGrammarEndingText(text: string): string {
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

    if (code >= KATAKANA_CODEPOINT_START && code <= KATAKANA_CODEPOINT_END) {
      normalized += String.fromCodePoint(code - KATAKANA_TO_HIRAGANA_OFFSET);
      continue;
    }

    normalized += char;
  }

  return normalized;
}

function matchesSuffix(text: string, suffixes: readonly string[]): boolean {
  return suffixes.some((suffix) => text === suffix);
}

function matchesPoliteCopulaEnding(text: string): boolean {
  if (!text.startsWith('です')) {
    return false;
  }

  return matchesSuffix(text.slice('です'.length), SENTENCE_FINAL_PARTICLE_SUFFIXES);
}

function matchesNegativeCopulaEnding(text: string): boolean {
  for (const prefix of NEGATIVE_COPULA_PREFIXES) {
    const negativeStem = `${prefix}ない`;
    if (!text.startsWith(negativeStem)) {
      continue;
    }

    const suffix = text.slice(negativeStem.length);
    return (
      matchesSuffix(suffix, SENTENCE_FINAL_PARTICLE_SUFFIXES) || matchesPoliteCopulaEnding(suffix)
    );
  }

  return false;
}

function matchesExplanatoryEnding(text: string): boolean {
  for (const prefix of EXPLANATORY_ENDING_PREFIXES) {
    if (EXPLANATORY_ENDING_THOUGHT_SUFFIXES.some((suffix) => text === `${prefix}${suffix}`)) {
      return true;
    }

    if (!text.startsWith(prefix)) {
      continue;
    }

    const suffix = text.slice(prefix.length);
    for (const core of EXPLANATORY_ENDING_CORES) {
      if (!suffix.startsWith(core)) {
        continue;
      }

      if (matchesSuffix(suffix.slice(core.length), EXPLANATORY_ENDING_TRAILING_PARTICLES)) {
        return true;
      }
    }
  }

  return false;
}

export function isStandaloneGrammarEndingText(text: string): boolean {
  const normalized = normalizeGrammarEndingText(text);
  if (!normalized) {
    return false;
  }

  return matchesPoliteCopulaEnding(normalized) || matchesNegativeCopulaEnding(normalized);
}

export function isSubtitleGrammarEndingText(text: string): boolean {
  const normalized = normalizeGrammarEndingText(text);
  if (!normalized) {
    return false;
  }

  return isStandaloneGrammarEndingText(normalized) || matchesExplanatoryEnding(normalized);
}
