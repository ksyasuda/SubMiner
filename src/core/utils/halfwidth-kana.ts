// No `u` flag: the range is entirely BMP so it changes nothing here, and
// Bun's unicode-mode matcher mis-handles this class next to certain ligatures.
const HALFWIDTH_KANA_RUN = /[｡-ﾟ]+/g;

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
  return text.replace(HALFWIDTH_KANA_RUN, (run) => run.normalize('NFKC'));
}
