import {
  JlptIgnoredPos1Entry,
  JLPT_IGNORED_MECAB_POS1,
  JLPT_IGNORED_MECAB_POS1_ENTRIES,
} from "./jlpt-ignored-mecab-pos1";

export { JLPT_IGNORED_MECAB_POS1_ENTRIES, JlptIgnoredPos1Entry };

// Data-driven MeCab POS names (pos1) used for JLPT filtering.
export const JLPT_IGNORED_MECAB_POS1_LIST: readonly string[] =
  JLPT_IGNORED_MECAB_POS1;

const JLPT_IGNORED_MECAB_POS1_SET = new Set<string>(
  JLPT_IGNORED_MECAB_POS1_LIST,
);

export function getIgnoredPos1Entries(): readonly JlptIgnoredPos1Entry[] {
  return JLPT_IGNORED_MECAB_POS1_ENTRIES;
}

export function shouldIgnoreJlptForMecabPos1(pos1: string): boolean {
  return JLPT_IGNORED_MECAB_POS1_SET.has(pos1);
}
