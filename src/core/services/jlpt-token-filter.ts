export type JlptIgnoredPos1Entry = {
  pos1: string;
  reason: string;
};

// Token-level lexical terms excluded from JLPT highlighting.
// These are not tied to POS and act as a safety layer for non-dictionary cases.
export const JLPT_EXCLUDED_TERMS = new Set([
  "この",
  "その",
  "あの",
  "どの",
  "これ",
  "それ",
  "あれ",
  "どれ",
  "ここ",
  "そこ",
  "あそこ",
  "どこ",
  "こと",
  "ああ",
  "ええ",
  "うう",
  "おお",
  "はは",
  "へえ",
  "ふう",
  "ほう",
]);

export function shouldIgnoreJlptByTerm(term: string): boolean {
  return JLPT_EXCLUDED_TERMS.has(term);
}

// MeCab POS1 categories that should be excluded from JLPT-level token tagging.
// These are filtered out because they are typically functional or non-lexical words.
export const JLPT_IGNORED_MECAB_POS1_ENTRIES = [
  {
    pos1: "助詞",
    reason: "Particles (ko/kara/nagara etc.): mostly grammatical glue, not independent vocabulary.",
  },
  {
    pos1: "助動詞",
    reason: "Auxiliary verbs (past tense, politeness, modality): grammar helpers.",
  },
  {
    pos1: "記号",
    reason: "Symbols/punctuation and symbols-like tokens.",
  },
  {
    pos1: "補助記号",
    reason: "Auxiliary symbols (e.g. bracket-like or markup tokens).",
  },
  {
    pos1: "連体詞",
    reason: "Adnominal forms (e.g. demonstratives like \"この\").",
  },
  {
    pos1: "感動詞",
    reason: "Interjections/onomatopoeia-style exclamations.",
  },
  {
    pos1: "接続詞",
    reason: "Conjunctions that connect clauses, usually not target vocab items.",
  },
  {
    pos1: "接頭詞",
    reason: "Prefixes/prefix-like grammatical elements.",
  },
] as const satisfies readonly JlptIgnoredPos1Entry[];

export const JLPT_IGNORED_MECAB_POS1 = JLPT_IGNORED_MECAB_POS1_ENTRIES.map(
  (entry) => entry.pos1,
);

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
