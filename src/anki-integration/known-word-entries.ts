// Known-word cache entries pair a word with the reading its Anki note teaches,
// so spelling collisions across readings (e.g. 床/ゆか vs 床/とこ) don't mark
// unrelated words as known. reading === null means the note carries no usable
// reading and the word matches in any reading (fail-open).
export interface KnownWordEntry {
  word: string;
  reading: string | null;
}

const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;
const FURIGANA_SEGMENT_PATTERN = /([^\s　\[\]]*)\[([^\]]*)\]/g;
const FURIGANA_BRACKET_PATTERN = /\[[^\]]*\]/g;
const WHITESPACE_PATTERN = /[\s　]+/g;

// Reading-bearing field names probed on every known-word note, in addition to
// any configured word fields (covers Kaishi's "Word Reading" and Lapis's
// "ExpressionReading" note types).
export const DEFAULT_KNOWN_WORD_READING_FIELDS = [
  'Reading',
  'Word Reading',
  'ExpressionReading',
  'Expression Reading',
];

export function isReadingFieldName(fieldName: string): boolean {
  return /reading/i.test(fieldName);
}

export function convertKatakanaToHiragana(text: string): string {
  let converted = '';
  for (const char of text) {
    const code = char.codePointAt(0);
    if (code !== undefined && code >= KATAKANA_CODEPOINT_START && code <= KATAKANA_CODEPOINT_END) {
      converted += String.fromCodePoint(code - KATAKANA_TO_HIRAGANA_OFFSET);
      continue;
    }
    converted += char;
  }
  return converted;
}

function isHiraganaReadingChar(char: string): boolean {
  const code = char.codePointAt(0);
  if (code === undefined) {
    return false;
  }
  return (code >= 0x3041 && code <= 0x309f) || code === 0x30fc;
}

// Splits Anki furigana syntax (`床[とこ]`, `お 決[き]まり`) into base text and
// reading. Values without brackets pass through with reading null.
export function parseFuriganaAnnotatedText(value: string): {
  text: string;
  reading: string | null;
} {
  if (!value.includes('[')) {
    return { text: value, reading: null };
  }
  const text = value.replace(FURIGANA_BRACKET_PATTERN, '').replace(WHITESPACE_PATTERN, '');
  const reading = value.replace(FURIGANA_SEGMENT_PATTERN, '$2').replace(WHITESPACE_PATTERN, '');
  return { text, reading: reading.length > 0 ? reading : null };
}

// Returns the hiragana-normalized reading, or '' when the value is not a
// plausible kana reading (callers fall back to text-only matching then).
export function normalizeKnownReadingForLookup(value: string): string {
  const parsed = parseFuriganaAnnotatedText(value.trim());
  const candidate = (parsed.reading ?? parsed.text).trim();
  if (!candidate) {
    return '';
  }
  const hiragana = convertKatakanaToHiragana(candidate);
  for (const char of hiragana) {
    if (!isHiraganaReadingChar(char)) {
      return '';
    }
  }
  return hiragana;
}

export function makeKnownWordEntryKey(entry: KnownWordEntry): string {
  return `${entry.word}\u0000${entry.reading ?? ''}`;
}

export function normalizeKnownWordEntryList(entries: KnownWordEntry[]): KnownWordEntry[] {
  const byKey = new Map<string, KnownWordEntry>();
  for (const entry of entries) {
    const word = entry.word.trim();
    if (!word) {
      continue;
    }
    const reading = entry.reading?.trim() || null;
    const normalized: KnownWordEntry = { word, reading };
    byKey.set(makeKnownWordEntryKey(normalized), normalized);
  }
  return [...byKey.values()].sort((left, right) =>
    makeKnownWordEntryKey(left).localeCompare(makeKnownWordEntryKey(right)),
  );
}

export function knownWordEntryListsEqual(left: KnownWordEntry[], right: KnownWordEntry[]): boolean {
  if (left.length !== right.length) {
    return false;
  }
  for (let index = 0; index < left.length; index += 1) {
    if (makeKnownWordEntryKey(left[index]!) !== makeKnownWordEntryKey(right[index]!)) {
      return false;
    }
  }
  return true;
}
