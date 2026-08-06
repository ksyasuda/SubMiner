// Kana classification and normalization for the injected scan runtime: the
// code-point ranges the walk tests every character against, and the folds that
// let halfwidth and katakana spellings compare equal to their dictionary form.
import { HAN_CODE_POINT_RANGES } from '../../text/han-code-points';

export const YOMITAN_KANA_HELPERS = String.raw`
      const HIRAGANA_CONVERSION_RANGE = [0x3041, 0x3096];
      const KATAKANA_CONVERSION_RANGE = [0x30a1, 0x30f6];
      const KANA_PROLONGED_SOUND_MARK_CODE_POINT = 0x30fc;
      const KATAKANA_SMALL_KA_CODE_POINT = 0x30f5;
      const KATAKANA_SMALL_KE_CODE_POINT = 0x30f6;
      const KANA_RANGES = [[0x3040, 0x309f], [0x30a0, 0x30ff], [0xff66, 0xff9f]];
      const HALFWIDTH_KATAKANA_RANGE = [0xff66, 0xff9d];
      const HALFWIDTH_KANA_PROLONGED_SOUND_MARK_CODE_POINT = 0xff70;
      // Folded one code point to one, so every index into a normalized string
      // still lines up with the original text — the name-candidate prefilter
      // and the furigana stem matching both index back into it. The standalone
      // voiced marks (ﾞ ﾟ) have no one-character equivalent and stay as they are.
      const HALFWIDTH_KATAKANA_TO_HIRAGANA = "をぁぃぅぇぉゃゅょっーあいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわん";
      function convertHalfwidthKanaCodePointToHiragana(codePoint) {
        if (codePoint < HALFWIDTH_KATAKANA_RANGE[0] || codePoint > HALFWIDTH_KATAKANA_RANGE[1]) { return null; }
        return HALFWIDTH_KATAKANA_TO_HIRAGANA[codePoint - HALFWIDTH_KATAKANA_RANGE[0]] || null;
      }
      // Halfwidth katakana is kana here but not to the rest of the pipeline
      // (known-word matching and frequency lookups only fold fullwidth), so a
      // reading taken from halfwidth text is written the way the fullwidth
      // katakana path already writes it. NFKC rather than the per-code-point
      // table: this is the one place where nothing indexes back into the
      // result, so a voiced pair (ｶ + ﾞ) can compose into the single ガ it
      // means instead of leaving a stray combining mark in the reading. Scoped
      // to the halfwidth runs, because NFKC over everything else rewrites
      // characters that have nothing to do with kana (① → 1, ㍑ → リットル).
      function convertHalfwidthKanaToKatakana(text) {
        return text.replace(/[ｦ-ﾟ]+/g, (run) => run.normalize("NFKC"));
      }
      // Han ranges come from the shared table so the scan walk and the character
      // dictionary agree on what a kanji is (supplementary planes included).
      // Halfwidth katakana counts as Japanese text: a name written that way has
      // to reach the greedy pre-pass, which has its own handling for it.
      const JAPANESE_RANGES = [[0x3040, 0x30ff], [0xff66, 0xff9f], ...${JSON.stringify(HAN_CODE_POINT_RANGES)}];
      function isCodePointInRange(codePoint, range) { return codePoint >= range[0] && codePoint <= range[1]; }
      function isCodePointInRanges(codePoint, ranges) { return ranges.some((range) => isCodePointInRange(codePoint, range)); }
      function isCodePointKana(codePoint) { return isCodePointInRanges(codePoint, KANA_RANGES); }
      function isCodePointJapanese(codePoint) { return isCodePointInRanges(codePoint, JAPANESE_RANGES); }
`;
