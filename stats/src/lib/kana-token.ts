const KANA_ONLY_TEXT = /^[\p{Script=Hiragana}\p{Script=Katakana}\u30fc\u309d\u309e\u30fd\u30fe]+$/u;

export function isKanaOnlyTokenText(text: string): boolean {
  const trimmed = text.trim();
  return trimmed.length > 0 && KANA_ONLY_TEXT.test(trimmed);
}
