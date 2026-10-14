// No `u` flag: the range is entirely BMP so it changes nothing here, and
// Bun's unicode-mode matcher mis-handles this class next to certain ligatures.
// The optional leading character is captured so a run opening with a voiced
// mark (カﾞ) can fold the mark into the fullwidth kana before it.
const HALFWIDTH_KANA_RUN = /([^｡-ﾟ]?)([｡-ﾟ]+)/g;
const LEADING_COMBINING_VOICED_MARK = /^[゙゚]/;

/**
 * Rewrites halfwidth katakana and its punctuation (ｱﾝﾀ, ﾀﾞｷﾆ, ｡｢｣､･) to the fullwidth
 * forms. Broadcast-derived subtitles write katakana halfwidth to save space; subtitle
 * text is normalized with this where it enters the app (live mpv text, parsed cues) so
 * display, lookups, mining, and stats all see one form.
 *
 * NFKC runs over the halfwidth runs only, never the whole string: it composes the voiced
 * pairs (ｶ + ﾞ) into single characters, but over everything it would also rewrite
 * unrelated text (① → 1, ！ → !, ㍑ → リットル).
 */
export function widenHalfwidthKana(text: string): string {
  return text.replace(HALFWIDTH_KANA_RUN, (_match, previous: string, run: string) => {
    const widened = run.normalize('NFKC');
    if (!previous || !LEADING_COMBINING_VOICED_MARK.test(widened)) {
      return previous + widened;
    }
    // A lone ﾞ widens to a combining mark; NFC composes it with the kana before it
    // (カ + ゛ → ガ) and leaves pairs that have no composed form untouched.
    return (previous + widened[0]).normalize('NFC') + widened.slice(1);
  });
}
