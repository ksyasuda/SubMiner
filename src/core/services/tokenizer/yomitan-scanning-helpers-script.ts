// Helper bundle for the in-page Yomitan scan runtime, composed from the
// fragments below. Injected as text into the parser window by
// yomitan-scan-runtime-script.ts, so it is data here, not code this process
// runs. The fragments are concatenated into a single function body and share
// one lexical scope: every function in them is hoisted, but the constants are
// not, so kana stays first — the later fragments read its ranges as they run.
import { YOMITAN_DICTIONARY_CLASSIFICATION_HELPERS } from './yomitan-dictionary-classification-script';
import { YOMITAN_FREQUENCY_HELPERS } from './yomitan-frequency-script';
import { YOMITAN_FURIGANA_HELPERS } from './yomitan-furigana-script';
import { YOMITAN_KANA_HELPERS } from './yomitan-kana-script';
import { YOMITAN_MATCH_SELECTION_HELPERS } from './yomitan-match-selection-script';

export { CHARACTER_DICTIONARY_TITLE_PREFIX } from './character-dictionary-title';

export const YOMITAN_SCANNING_HELPERS = [
  YOMITAN_KANA_HELPERS,
  YOMITAN_FURIGANA_HELPERS,
  YOMITAN_FREQUENCY_HELPERS,
  YOMITAN_DICTIONARY_CLASSIFICATION_HELPERS,
  YOMITAN_MATCH_SELECTION_HELPERS,
].join('\n');
