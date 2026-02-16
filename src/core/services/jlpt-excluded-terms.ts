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
