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

// Inverse of widenHalfwidthKana per character, derived from it so the two stay in sync.
// Voiced pairs (ｶﾞ) are included because they widen to a single character (ガ).
const HALFWIDTH_BY_FULLWIDTH = new Map<string, string>();
for (let code = 0xff61; code <= 0xff9f; code += 1) {
  const halfwidth = String.fromCharCode(code);
  for (const form of [halfwidth, `${halfwidth}ﾞ`, `${halfwidth}ﾟ`]) {
    const widened = widenHalfwidthKana(form);
    if (widened.length === 1 && !HALFWIDTH_BY_FULLWIDTH.has(widened)) {
      HALFWIDTH_BY_FULLWIDTH.set(widened, form);
    }
  }
}

/**
 * Rewrites fullwidth katakana and punctuation to halfwidth forms (アンタ → ｱﾝﾀ). Used to
 * search text stored before subtitle ingest widened it; display paths should widen instead.
 */
export function narrowFullwidthKana(text: string): string {
  return Array.from(text, (char) => HALFWIDTH_BY_FULLWIDTH.get(char) ?? char).join('');
}
