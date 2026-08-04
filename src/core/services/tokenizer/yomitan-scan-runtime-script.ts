// In-page Yomitan scan runtime: the helper bundle and scan walk that get
// installed once per parser window as globalThis.__subminerYomitanScan, plus
// the tiny per-line call script. Kept separate from the host runtime module so
// the injected-script text (which is data, not executed here) does not
// dominate that file.

export type YomitanFrequencyMode = 'occurrence-based' | 'rank-based';

export const CHARACTER_DICTIONARY_TITLE_PREFIX = 'SubMiner Character Dictionary';

const YOMITAN_SCANNING_HELPERS = String.raw`
      const HIRAGANA_CONVERSION_RANGE = [0x3041, 0x3096];
      const KATAKANA_CONVERSION_RANGE = [0x30a1, 0x30f6];
      const KANA_PROLONGED_SOUND_MARK_CODE_POINT = 0x30fc;
      const KATAKANA_SMALL_KA_CODE_POINT = 0x30f5;
      const KATAKANA_SMALL_KE_CODE_POINT = 0x30f6;
      const KANA_RANGES = [[0x3040, 0x309f], [0x30a0, 0x30ff]];
      const JAPANESE_RANGES = [[0x3040, 0x30ff], [0x3400, 0x9fff]];
      function isCodePointInRange(codePoint, range) { return codePoint >= range[0] && codePoint <= range[1]; }
      function isCodePointInRanges(codePoint, ranges) { return ranges.some((range) => isCodePointInRange(codePoint, range)); }
      function isCodePointKana(codePoint) { return isCodePointInRanges(codePoint, KANA_RANGES); }
      function isCodePointJapanese(codePoint) { return isCodePointInRanges(codePoint, JAPANESE_RANGES); }
      function createFuriganaSegment(text, reading) { return {text, reading}; }
      function getSegmentReadingContribution(segment) {
        if (typeof segment.reading === "string" && segment.reading.length > 0) { return segment.reading; }
        const segmentText = typeof segment.text === "string" ? segment.text : "";
        const isKanaOnly = segmentText.length > 0 && [...segmentText].every((char) => isCodePointKana(char.codePointAt(0)));
        return isKanaOnly ? segmentText : "";
      }
      function getProlongedHiragana(previousCharacter) {
        switch (previousCharacter) {
          case "あ": case "か": case "が": case "さ": case "ざ": case "た": case "だ": case "な": case "は": case "ば": case "ぱ": case "ま": case "や": case "ら": case "わ": case "ぁ": case "ゃ": case "ゎ": return "あ";
          case "い": case "き": case "ぎ": case "し": case "じ": case "ち": case "ぢ": case "に": case "ひ": case "び": case "ぴ": case "み": case "り": case "ぃ": return "い";
          case "う": case "く": case "ぐ": case "す": case "ず": case "つ": case "づ": case "ぬ": case "ふ": case "ぶ": case "ぷ": case "む": case "ゆ": case "る": case "ぅ": case "ゅ": return "う";
          case "え": case "け": case "げ": case "せ": case "ぜ": case "て": case "で": case "ね": case "へ": case "べ": case "ぺ": case "め": case "れ": case "ぇ": return "え";
          case "お": case "こ": case "ご": case "そ": case "ぞ": case "と": case "ど": case "の": case "ほ": case "ぼ": case "ぽ": case "も": case "よ": case "ろ": case "を": case "ぉ": case "ょ": return "う";
          default: return null;
        }
      }
      function getFuriganaKanaSegments(text, reading) {
        const newSegments = [];
        let start = 0;
        let state = (reading[0] === text[0]);
        for (let i = 1; i < text.length; ++i) {
          const newState = (reading[i] === text[i]);
          if (state === newState) { continue; }
          newSegments.push(createFuriganaSegment(text.substring(start, i), state ? '' : reading.substring(start, i)));
          state = newState;
          start = i;
        }
        newSegments.push(createFuriganaSegment(text.substring(start), state ? '' : reading.substring(start)));
        return newSegments;
      }
      function convertKatakanaToHiragana(text, keepProlongedSoundMarks = false) {
        let result = '';
        const offset = (HIRAGANA_CONVERSION_RANGE[0] - KATAKANA_CONVERSION_RANGE[0]);
        for (let char of text) {
          const codePoint = char.codePointAt(0);
          switch (codePoint) {
            case KATAKANA_SMALL_KA_CODE_POINT:
            case KATAKANA_SMALL_KE_CODE_POINT:
              break;
            case KANA_PROLONGED_SOUND_MARK_CODE_POINT:
              if (!keepProlongedSoundMarks && result.length > 0) {
                const char2 = getProlongedHiragana(result[result.length - 1]);
                if (char2 !== null) { char = char2; }
              }
              break;
            default:
              if (isCodePointInRange(codePoint, KATAKANA_CONVERSION_RANGE)) {
                char = String.fromCodePoint(codePoint + offset);
              }
              break;
          }
          result += char;
        }
        return result;
      }
      function segmentizeFurigana(reading, readingNormalized, groups, groupsStart) {
        const groupCount = groups.length - groupsStart;
        if (groupCount <= 0) { return reading.length === 0 ? [] : null; }
        const group = groups[groupsStart];
        const {isKana, text} = group;
        if (isKana) {
          if (group.textNormalized !== null && readingNormalized.startsWith(group.textNormalized)) {
            const segments = segmentizeFurigana(reading.substring(text.length), readingNormalized.substring(text.length), groups, groupsStart + 1);
            if (segments !== null) {
              if (reading.startsWith(text)) { segments.unshift(createFuriganaSegment(text, '')); }
              else { segments.unshift(...getFuriganaKanaSegments(text, reading)); }
              return segments;
            }
          }
          return null;
        }
        let result = null;
        for (let i = reading.length; i >= text.length; --i) {
          const segments = segmentizeFurigana(reading.substring(i), readingNormalized.substring(i), groups, groupsStart + 1);
          if (segments !== null) {
            if (result !== null) { return null; }
            segments.unshift(createFuriganaSegment(text, reading.substring(0, i)));
            result = segments;
          }
          if (groupCount === 1) { break; }
        }
        return result;
      }
      function distributeFurigana(term, reading) {
        if (reading === term) { return [createFuriganaSegment(term, '')]; }
        const groups = [];
        let groupPre = null;
        let isKanaPre = null;
        for (const c of term) {
          const isKana = isCodePointKana(c.codePointAt(0));
          if (isKana === isKanaPre) { groupPre.text += c; }
          else {
            groupPre = {isKana, text: c, textNormalized: null};
            groups.push(groupPre);
            isKanaPre = isKana;
          }
        }
        for (const group of groups) {
          if (group.isKana) { group.textNormalized = convertKatakanaToHiragana(group.text); }
        }
        const segments = segmentizeFurigana(reading, convertKatakanaToHiragana(reading), groups, 0);
        return segments !== null ? segments : [createFuriganaSegment(term, reading)];
      }
      function getStemLength(text1, text2) {
        const minLength = Math.min(text1.length, text2.length);
        if (minLength === 0) { return 0; }
        let i = 0;
        while (true) {
          const char1 = text1.codePointAt(i);
          const char2 = text2.codePointAt(i);
          if (char1 !== char2) { break; }
          const charLength = String.fromCodePoint(char1).length;
          i += charLength;
          if (i >= minLength) {
            if (i > minLength) { i -= charLength; }
            break;
          }
        }
        return i;
      }
      function distributeFuriganaInflected(term, reading, source) {
        const termNormalized = convertKatakanaToHiragana(term);
        const readingNormalized = convertKatakanaToHiragana(reading);
        const sourceNormalized = convertKatakanaToHiragana(source);
        let mainText = term;
        let stemLength = getStemLength(termNormalized, sourceNormalized);
        const readingStemLength = getStemLength(readingNormalized, sourceNormalized);
        if (readingStemLength > 0 && readingStemLength >= stemLength) {
          mainText = reading;
          stemLength = readingStemLength;
          reading = source.substring(0, stemLength) + reading.substring(stemLength);
        }
        const segments = [];
        if (stemLength > 0) {
          mainText = source.substring(0, stemLength) + mainText.substring(stemLength);
          const segments2 = distributeFurigana(mainText, reading);
          let consumed = 0;
          for (const segment of segments2) {
            const start = consumed;
            consumed += segment.text.length;
            if (consumed < stemLength) { segments.push(segment); }
            else if (consumed === stemLength) { segments.push(segment); break; }
            else {
              if (start < stemLength) { segments.push(createFuriganaSegment(mainText.substring(start, stemLength), '')); }
              break;
            }
          }
        }
        if (stemLength < source.length) {
          const remainder = source.substring(stemLength);
          const last = segments[segments.length - 1];
          if (last && last.reading.length === 0) { last.text += remainder; }
          else { segments.push(createFuriganaSegment(remainder, '')); }
        }
        return segments;
      }
      function parsePositiveFrequencyNumber(value) {
        if (typeof value === 'number' && Number.isFinite(value) && value > 0) {
          return Math.max(1, Math.floor(value));
        }
        if (typeof value === 'string') {
          const numericMatch = value.trim().match(/[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/)?.[0];
          if (!numericMatch) { return null; }
          const parsed = Number.parseFloat(numericMatch);
          if (!Number.isFinite(parsed) || parsed <= 0) { return null; }
          return Math.max(1, Math.floor(parsed));
        }
        if (Array.isArray(value)) {
          for (const item of value) {
            const parsed = parsePositiveFrequencyNumber(item);
            if (parsed !== null) { return parsed; }
          }
        }
        return null;
      }
      function parseDisplayFrequencyNumber(value) {
        if (typeof value === 'string') {
          const leadingDigits = value.trim().match(/^\d+/)?.[0];
          if (!leadingDigits) { return null; }
          const parsed = Number.parseInt(leadingDigits, 10);
          return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
        }
        return parsePositiveFrequencyNumber(value);
      }
      function getFrequencyDictionaryName(frequency) {
        const candidates = [
          frequency?.dictionary,
          frequency?.dictionaryName,
          frequency?.name,
          frequency?.title,
          frequency?.dictionaryTitle,
          frequency?.dictionaryAlias
        ];
        for (const candidate of candidates) {
          if (typeof candidate === 'string' && candidate.trim().length > 0) {
            return candidate.trim();
          }
        }
        return null;
      }
      function getBestFrequencyRank(dictionaryEntry, headwordIndex, dictionaryPriorityByName, dictionaryFrequencyModeByName) {
        let best = null;
        const headwordCount = Array.isArray(dictionaryEntry?.headwords) ? dictionaryEntry.headwords.length : 0;
        for (const frequency of dictionaryEntry?.frequencies || []) {
          if (!frequency || typeof frequency !== 'object') { continue; }
          const frequencyHeadwordIndex = frequency.headwordIndex;
          if (typeof frequencyHeadwordIndex === 'number') {
            if (frequencyHeadwordIndex !== headwordIndex) { continue; }
          } else if (headwordCount > 1) {
            continue;
          }
          const dictionary = getFrequencyDictionaryName(frequency);
          if (!dictionary) { continue; }
          if (dictionaryFrequencyModeByName[dictionary] === 'occurrence-based') { continue; }
          const rank =
            parseDisplayFrequencyNumber(frequency.displayValue) ??
            parsePositiveFrequencyNumber(frequency.frequency);
          if (rank === null) { continue; }
          const priorityRaw = dictionaryPriorityByName[dictionary];
          const fallbackPriority =
            typeof frequency.dictionaryIndex === 'number' && Number.isFinite(frequency.dictionaryIndex)
              ? Math.max(0, Math.floor(frequency.dictionaryIndex))
              : Number.MAX_SAFE_INTEGER;
          const priority =
            typeof priorityRaw === 'number' && Number.isFinite(priorityRaw)
              ? Math.max(0, Math.floor(priorityRaw))
              : fallbackPriority;
          if (best === null || priority < best.priority || (priority === best.priority && rank < best.rank)) {
            best = { priority, rank };
          }
        }
        return best?.rank ?? null;
      }
      function hasExactSource(headword, token, requirePrimary) {
        for (const src of headword.sources || []) {
          if (src.originalText !== token) { continue; }
          if (requirePrimary && !src.isPrimary) { continue; }
          if (src.matchType !== 'exact') { continue; }
          return true;
        }
        return false;
      }
      function collectExactHeadwordMatches(dictionaryEntries, token, requirePrimary) {
        const matches = [];
        for (const dictionaryEntry of dictionaryEntries || []) {
          const headwords = Array.isArray(dictionaryEntry?.headwords) ? dictionaryEntry.headwords : [];
          for (let headwordIndex = 0; headwordIndex < headwords.length; headwordIndex += 1) {
            const headword = headwords[headwordIndex];
            if (!hasExactSource(headword, token, requirePrimary)) { continue; }
            matches.push({ dictionaryEntry, headword, headwordIndex });
          }
        }
        return matches;
      }
      function sameHeadword(match, preferredMatch) {
        if (!match || !preferredMatch) {
          return false;
        }
        if (match.headword?.term !== preferredMatch.headword?.term) {
          return false;
        }
        const matchReading = typeof match.headword?.reading === 'string' ? match.headword.reading : '';
        const preferredReading =
          typeof preferredMatch.headword?.reading === 'string' ? preferredMatch.headword.reading : '';
        if (!matchReading || !preferredReading) {
          return true;
        }
        return matchReading === preferredReading;
      }
      function getBestFrequencyRankForMatches(matches, dictionaryPriorityByName, dictionaryFrequencyModeByName) {
        let best = null;
        for (const match of matches) {
          const rank = getBestFrequencyRank(
            match.dictionaryEntry,
            match.headwordIndex,
            dictionaryPriorityByName,
            dictionaryFrequencyModeByName
          );
          if (rank === null) { continue; }
          if (best === null || rank < best) {
            best = rank;
          }
        }
        return best;
      }
      function normalizeWordClasses(headword) {
          if (!Array.isArray(headword?.wordClasses)) { return undefined; }
          const classes = headword.wordClasses.filter((wordClass) => typeof wordClass === "string" && wordClass.trim().length > 0);
          return classes.length > 0 ? classes : undefined;
        }
        function appendDictionaryNames(target, value) {
          if (!value || typeof value !== 'object') {
            return;
          }
          const candidates = [
            value.dictionary,
            value.dictionaryName,
            value.name,
            value.title,
            value.dictionaryTitle,
            value.dictionaryAlias
          ];
          for (const candidate of candidates) {
            if (typeof candidate === 'string' && candidate.trim().length > 0) {
              target.push(candidate.trim());
            }
          }
        }
        function getDictionaryEntryNames(entry) {
          const names = [];
          appendDictionaryNames(names, entry);
          for (const definition of entry?.definitions || []) {
            appendDictionaryNames(names, definition);
          }
          for (const frequency of entry?.frequencies || []) {
            appendDictionaryNames(names, frequency);
          }
          for (const pronunciation of entry?.pronunciations || []) {
            appendDictionaryNames(names, pronunciation);
          }
          return names;
        }
        function isNameDictionaryEntry(entry) {
          if (!includeNameMatchMetadata || !entry || typeof entry !== 'object') {
            return false;
          }
          return getDictionaryEntryNames(entry).some((name) => name.startsWith(${JSON.stringify(CHARACTER_DICTIONARY_TITLE_PREFIX)}));
        }
        function parseSubMinerMediaIdFromString(value) {
          const imageMatch = value.match(/\bimg\/m(\d+)-/i);
          if (imageMatch) {
            const parsed = Number.parseInt(imageMatch[1], 10);
            if (Number.isSafeInteger(parsed) && parsed > 0) { return parsed; }
          }
          const titleMatch = value.match(/${CHARACTER_DICTIONARY_TITLE_PREFIX}[^\d]*(?:AniList\s*)?(\d+)/i);
          if (titleMatch) {
            const parsed = Number.parseInt(titleMatch[1], 10);
            if (Number.isSafeInteger(parsed) && parsed > 0) { return parsed; }
          }
          return null;
        }
        function parseSubMinerMediaIdCandidate(value) {
          if (typeof value === 'number' && Number.isSafeInteger(value) && value > 0) {
            return value;
          }
          if (typeof value === 'string' && /^\d+$/.test(value.trim())) {
            const parsed = Number.parseInt(value.trim(), 10);
            if (Number.isSafeInteger(parsed) && parsed > 0) { return parsed; }
          }
          return null;
        }
        function collectSubMinerMediaIds(value, target) {
          if (typeof value === 'string') {
            const parsed = parseSubMinerMediaIdFromString(value);
            if (parsed !== null) { target.add(parsed); }
            return;
          }
          if (!value || typeof value !== 'object') {
            return;
          }
          if (Array.isArray(value)) {
            for (const item of value) { collectSubMinerMediaIds(item, target); }
            return;
          }
          const mediaIdCandidates = [
            value.subminerMediaId,
            value.subMinerMediaId,
            value.characterDictionaryMediaId,
            value.data?.subminerMediaId,
            value.data?.subMinerMediaId,
            value.data?.characterDictionaryMediaId
          ];
          for (const candidate of mediaIdCandidates) {
            const parsed = parseSubMinerMediaIdCandidate(candidate);
            if (parsed !== null) { target.add(parsed); }
          }
          for (const child of Object.values(value)) {
            collectSubMinerMediaIds(child, target);
          }
        }
        function getSubMinerMediaIds(entry) {
          const mediaIds = new Set();
          collectSubMinerMediaIds(entry, mediaIds);
          return mediaIds;
        }
        function isCurrentMediaNameDictionaryEntry(entry) {
          if (!isNameDictionaryEntry(entry)) {
            return false;
          }
          if (currentCharacterDictionaryMediaId === null) {
            return true;
          }
          const mediaIds = getSubMinerMediaIds(entry);
          return mediaIds.size === 0 || mediaIds.has(currentCharacterDictionaryMediaId);
      }
      function findLongestNameMatch(dictionaryEntries, textWindow) {
        let best = null;
        for (const dictionaryEntry of dictionaryEntries || []) {
          if (!isCurrentMediaNameDictionaryEntry(dictionaryEntry)) { continue; }
          const headwords = Array.isArray(dictionaryEntry?.headwords) ? dictionaryEntry.headwords : [];
          for (let headwordIndex = 0; headwordIndex < headwords.length; headwordIndex += 1) {
            const headword = headwords[headwordIndex];
            for (const src of headword?.sources || []) {
              if (src.matchType !== 'exact' || src.isPrimary !== true) { continue; }
              const originalText = typeof src.originalText === 'string' ? src.originalText : '';
              if (!originalText || !textWindow.startsWith(originalText)) { continue; }
              if (best === null || originalText.length > best.sourceLength) {
                best = { dictionaryEntry, headword, headwordIndex, sourceLength: originalText.length };
              }
            }
          }
        }
        return best;
      }
      function findLongestGenericMatchLength(dictionaryEntries, textWindow) {
        let best = 0;
        for (const dictionaryEntry of dictionaryEntries || []) {
          if (isNameDictionaryEntry(dictionaryEntry)) { continue; }
          const headwords = Array.isArray(dictionaryEntry?.headwords) ? dictionaryEntry.headwords : [];
          for (const headword of headwords) {
            for (const src of headword?.sources || []) {
              if (src.matchType !== 'exact' || src.isPrimary !== true) { continue; }
              const originalText = typeof src.originalText === 'string' ? src.originalText : '';
              if (!originalText || !textWindow.startsWith(originalText)) { continue; }
              if (originalText.length > best) { best = originalText.length; }
            }
          }
        }
        return best;
      }
      function getPreferredHeadword(dictionaryEntries, token, dictionaryPriorityByName, dictionaryFrequencyModeByName) {
        const currentMediaDictionaryEntries =
          currentCharacterDictionaryMediaId === null
            ? (dictionaryEntries || [])
            : (dictionaryEntries || []).filter((entry) => {
                if (!isNameDictionaryEntry(entry)) { return true; }
                return isCurrentMediaNameDictionaryEntry(entry);
              });
        const exactPrimaryMatches = collectExactHeadwordMatches(currentMediaDictionaryEntries, token, true);
        let matchedNameDictionary = false;
        if (includeNameMatchMetadata) {
          for (const dictionaryEntry of currentMediaDictionaryEntries || []) {
            if (!isCurrentMediaNameDictionaryEntry(dictionaryEntry)) { continue; }
            for (const match of exactPrimaryMatches) {
              if (match.dictionaryEntry !== dictionaryEntry) { continue; }
              matchedNameDictionary = true;
              break;
            }
            if (matchedNameDictionary) { break; }
          }
        }
        const preferredMatch = exactPrimaryMatches[0];
        if (preferredMatch) {
          const exactFrequencyMatches = collectExactHeadwordMatches(currentMediaDictionaryEntries, token, false)
            .filter((match) => sameHeadword(match, preferredMatch));
          return {
            term: preferredMatch.headword.term,
            reading: preferredMatch.headword.reading,
            wordClasses: normalizeWordClasses(preferredMatch.headword),
            isNameMatch:
              matchedNameDictionary || isCurrentMediaNameDictionaryEntry(preferredMatch.dictionaryEntry),
            frequencyRank: getBestFrequencyRankForMatches(
              exactFrequencyMatches.length > 0 ? exactFrequencyMatches : exactPrimaryMatches,
              dictionaryPriorityByName,
              dictionaryFrequencyModeByName
            )
          };
        }
        return null;
      }
`;

// Bump whenever the install script below changes so already-loaded parser
// windows re-install the new scan runtime instead of running the stale one.
export const YOMITAN_SCAN_RUNTIME_VERSION = 4;
export const YOMITAN_SCAN_RUNTIME_MISSING_SENTINEL = '__subminer-yomitan-scan-runtime-missing__';

export interface YomitanScanRequestParams {
  text: string;
  profileIndex: number;
  scanLength: number;
  includeNameMatchMetadata: boolean;
  greedyNameScanEnabled: boolean;
  currentCharacterDictionaryMediaId: number | null;
  dictionaryPriorityByName: Record<string, number>;
  dictionaryFrequencyModeByName: Partial<Record<string, YomitanFrequencyMode>>;
  cacheEpoch: number;
  /**
   * Key of the character-name candidate list installed for the current media,
   * or null to scan every Japanese position (see the pre-pass prefilter).
   */
  nameCandidateKey: string | null;
}

// Installed once per parser window (and re-installed after in-page reloads):
// keeps V8 from re-parsing the helper bundle on every subtitle line, and hosts
// the cross-line termsFind cache. Each subtitle line then only evaluates a tiny
// call into globalThis.__subminerYomitanScan.
export const YOMITAN_SCAN_RUNTIME_INSTALL_SCRIPT = String.raw`
  (() => {
    if (globalThis.__subminerYomitanScanVersion === ${YOMITAN_SCAN_RUNTIME_VERSION}) {
      return true;
    }
    const invoke = (action, params) =>
      new Promise((resolve, reject) => {
        chrome.runtime.sendMessage({ action, params }, (response) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
            return;
          }
          if (!response || typeof response !== "object") {
            reject(new Error("Invalid response from Yomitan backend"));
            return;
          }
          if (response.error) {
            reject(new Error(response.error.message || "Yomitan backend error"));
            return;
          }
          resolve(response.result);
        });
      });
    // Cross-line termsFind LRU keyed by profile + substring: subtitle lines
    // repeat particles and inflections constantly, so most lookups hit here.
    // Entries hold in-flight promises so concurrent identical lookups dedupe.
    const termsFindCache = new Map();
    const TERMS_FIND_CACHE_LIMIT = 2000;
    let termsFindCacheEpoch = -1;
    const MAX_SHRINKING_WINDOW_RETRY_LOOKUPS = 4;
    // Character-name candidate forms for the current media, installed
    // separately from the per-line scan call so the per-line script stays tiny.
    // Stored raw here; the normalized lookup index is built inside the scan,
    // where the kana-normalization helper is in scope, and reused by key.
    let rawNameCandidates = null;
    let nameCandidateIndex = null;
    globalThis.__subminerYomitanScanSetNameCandidates = (key, forms) => {
      if (!key || !Array.isArray(forms) || forms.length === 0) {
        rawNameCandidates = null;
        nameCandidateIndex = null;
        return false;
      }
      rawNameCandidates = { key, forms };
      nameCandidateIndex = null;
      return true;
    };
    globalThis.__subminerYomitanScanVersion = ${YOMITAN_SCAN_RUNTIME_VERSION};
    globalThis.__subminerYomitanScan = async (scanParams) => {
      const {
        text,
        profileIndex,
        scanLength,
        includeNameMatchMetadata,
        greedyNameScanEnabled,
        currentCharacterDictionaryMediaId,
        dictionaryPriorityByName,
        dictionaryFrequencyModeByName,
        cacheEpoch,
        nameCandidateKey
      } = scanParams;
      if (cacheEpoch !== termsFindCacheEpoch) {
        termsFindCache.clear();
        termsFindCacheEpoch = cacheEpoch;
      }
${YOMITAN_SCANNING_HELPERS}
      const CAPTION_OPENING_BRACKETS = new Set(["(", "（", "[", "［", "{", "｛", "「", "『", "【", "〈", "《", "≪", "＜", "<"]);
      function shouldEmitUnparsedRunAsToken(runText) {
        if (!/[\p{L}\p{N}]/u.test(runText)) { return false; }
        const firstChar = Array.from(runText.trim())[0];
        return firstChar !== undefined && !CAPTION_OPENING_BRACKETS.has(firstChar);
      }
      function isLookupWorthyCodePoint(codePoint) {
        if (isCodePointJapanese(codePoint)) { return true; }
        return /[\p{L}\p{N}]/u.test(String.fromCodePoint(codePoint));
      }
      function isKanaOnlyRunText(runText) {
        const chars = Array.from(runText);
        return chars.length > 0 && chars.every((char) => isCodePointKana(char.codePointAt(0)));
      }
      const details = {matchType: "exact", deinflect: true};
      const tokens = [];
      async function termsFindAt(position, windowLength) {
        const substring = text.substring(position, position + windowLength);
        const cacheKey = profileIndex + " " + substring;
        const cached = termsFindCache.get(cacheKey);
        if (cached !== undefined) {
          termsFindCache.delete(cacheKey);
          termsFindCache.set(cacheKey, cached);
          return await cached;
        }
        const pending = invoke("termsFind", { text: substring, details, optionsContext: { index: profileIndex } });
        termsFindCache.set(cacheKey, pending);
        while (termsFindCache.size > TERMS_FIND_CACHE_LIMIT) {
          const oldestKey = termsFindCache.keys().next().value;
          if (oldestKey === undefined) { break; }
          termsFindCache.delete(oldestKey);
        }
        try {
          return await pending;
        } catch (error) {
          termsFindCache.delete(cacheKey);
          throw error;
        }
      }
      // Text the walk skips accumulates into unparsed runs, mirroring the
      // filler chunks the parseText segmentation used to provide: runs stay
      // hoverable (flagged isUnparsedRun) unless they are punctuation-only or
      // caption-style asides, and kana continuations of a longer headword
      // extend the previous token instead.
      function flushUnparsedRun(runStart, runEnd) {
        if (runStart === null || runEnd <= runStart) { return; }
        const runText = text.substring(runStart, runEnd);
        const previousToken = tokens[tokens.length - 1];
        if (
          previousToken &&
          previousToken.endPos === runStart &&
          isKanaOnlyRunText(runText) &&
          typeof previousToken.headword === "string" &&
          previousToken.headword.length > previousToken.surface.length &&
          previousToken.headword.startsWith(previousToken.surface + runText)
        ) {
          previousToken.surface += runText;
          // The run is kana-only, so its reading is itself: append it or the
          // reading stops covering the surface, which disables the known-word
          // reading fallback (isCompleteReadingForSurface) downstream.
          previousToken.reading += runText;
          // The run is kana-only, so its reading is itself: append it or the
          // reading stops covering the surface, which disables the known-word
          // reading fallback (isCompleteReadingForSurface) downstream.
          previousToken.endPos = runEnd;
          return;
        }
        if (!shouldEmitUnparsedRunAsToken(runText)) { return; }
        tokens.push({
          surface: runText,
          reading: "",
          headword: runText,
          startPos: runStart,
          endPos: runEnd,
          isUnparsedRun: true
        });
      }
      function buildScanToken(position, source, preferredHeadword) {
        const reading = typeof preferredHeadword.reading === "string" ? preferredHeadword.reading : "";
        const segments = distributeFuriganaInflected(preferredHeadword.term, reading, source);
        const tokenPayload = {
          surface: segments.map((segment) => segment.text).join("") || source,
          reading: segments.map(getSegmentReadingContribution).join(""),
          headword: preferredHeadword.term,
          headwordReading: reading || undefined,
          startPos: position,
          endPos: position + source.length,
          isNameMatch: includeNameMatchMetadata && preferredHeadword.isNameMatch === true,
          frequencyRank:
            typeof preferredHeadword.frequencyRank === "number" && Number.isFinite(preferredHeadword.frequencyRank)
              ? Math.max(1, Math.floor(preferredHeadword.frequencyRank))
              : undefined,
        };
        if (Array.isArray(preferredHeadword.wordClasses) && preferredHeadword.wordClasses.length > 0) {
          tokenPayload.wordClasses = preferredHeadword.wordClasses;
        }
        return tokenPayload;
      }
      async function findTokenAt(position, windowLength) {
        const codePoint = text.codePointAt(position);
        const character = String.fromCodePoint(codePoint);
        const result = await termsFindAt(position, windowLength);
        const dictionaryEntries = Array.isArray(result?.dictionaryEntries) ? result.dictionaryEntries : [];
        const originalTextLength = typeof result?.originalTextLength === "number" ? result.originalTextLength : 0;
        if (dictionaryEntries.length === 0 || originalTextLength <= 0 || (originalTextLength === character.length && !isCodePointJapanese(codePoint))) {
          return { token: null, matchedLength: 0 };
        }
        const source = text.substring(position, position + originalTextLength);
        const preferredHeadword = getPreferredHeadword(
          dictionaryEntries,
          source,
          dictionaryPriorityByName,
          dictionaryFrequencyModeByName
        );
        if (!preferredHeadword || typeof preferredHeadword.term !== "string") {
          return { token: null, matchedLength: originalTextLength };
        }
        return { token: buildScanToken(position, source, preferredHeadword), matchedLength: originalTextLength };
      }
      // Halfwidth katakana survives kana normalization unchanged, so a name
      // written that way would not prefix-match a candidate form. Those
      // positions bypass the prefilter rather than risk a missed name.
      function isHalfwidthKatakanaCodePoint(codePoint) {
        return codePoint >= 0xff66 && codePoint <= 0xff9f;
      }
      // Build (once per candidate list) a first-character bucket index of the
      // normalized name forms, so the pre-pass can reject a position with a
      // single map hit instead of a backend round trip.
      if (rawNameCandidates && nameCandidateIndex?.key !== rawNameCandidates.key) {
        const byFirstChar = new Map();
        for (const form of rawNameCandidates.forms) {
          const normalized = typeof form === "string" ? convertKatakanaToHiragana(form.trim()) : "";
          if (!normalized) { continue; }
          const bucket = byFirstChar.get(normalized[0]);
          if (bucket) { bucket.push(normalized); } else { byFirstChar.set(normalized[0], [normalized]); }
        }
        nameCandidateIndex = byFirstChar.size > 0 ? { key: rawNameCandidates.key, byFirstChar } : null;
      } else if (!rawNameCandidates) {
        nameCandidateIndex = null;
      }
      // Only meaningful when the installed list matches the media this scan is
      // for; otherwise fall back to scanning every position.
      const activeNameCandidateIndex =
        nameCandidateKey !== null && nameCandidateIndex?.key === nameCandidateKey
          ? nameCandidateIndex
          : null;
      const normalizedText = activeNameCandidateIndex ? convertKatakanaToHiragana(text) : "";
      // Yomitan collapses emphatic sequences before matching (すっっごーーい →
      // すごい), so a stretched name still resolves to its entry. Skipping these
      // characters keeps such spellings candidates; the filter only ever grows
      // the probe set, so a false positive costs one lookup, never a name.
      const EMPHATIC_SKIP_CHARS = new Set(["ぁ", "ぃ", "ぅ", "ぇ", "ぉ", "っ", "ゃ", "ゅ", "ょ", "ー"]);
      function matchesCandidateFormAt(form, position) {
        let textIndex = position;
        for (let formIndex = 0; formIndex < form.length; formIndex += 1) {
          while (
            textIndex < normalizedText.length &&
            normalizedText[textIndex] !== form[formIndex] &&
            EMPHATIC_SKIP_CHARS.has(normalizedText[textIndex])
          ) {
            textIndex += 1;
          }
          if (normalizedText[textIndex] !== form[formIndex]) { return false; }
          textIndex += 1;
        }
        return true;
      }
      function couldNameStartAt(position, codePoint) {
        if (!activeNameCandidateIndex) { return true; }
        if (isHalfwidthKatakanaCodePoint(codePoint)) { return true; }
        const bucket = activeNameCandidateIndex.byFirstChar.get(normalizedText[position]);
        if (!bucket) { return false; }
        for (const form of bucket) {
          if (matchesCandidateFormAt(form, position)) { return true; }
        }
        return false;
      }
      // Greedy name pre-pass: character-name matches claim their spans before
      // the left-to-right walk, so a longer generic match starting earlier
      // (e.g. とヨー → 渡洋) cannot swallow the start of a name (ヨータ).
      const nameTokens = [];
      if (greedyNameScanEnabled) {
        let namePos = 0;
        while (namePos < text.length) {
          const codePoint = text.codePointAt(namePos);
          if (!isCodePointJapanese(codePoint) || !couldNameStartAt(namePos, codePoint)) {
            namePos += String.fromCodePoint(codePoint).length;
            continue;
          }
          const result = await termsFindAt(namePos, scanLength);
          const dictionaryEntries = Array.isArray(result?.dictionaryEntries) ? result.dictionaryEntries : [];
          const textWindow = text.substring(namePos, namePos + scanLength);
          const nameMatch = findLongestNameMatch(dictionaryEntries, textWindow);
          // A name only claims its span when no strictly longer generic word
          // starts at the same position (a character named 空 must not split
          // 空気). Ties go to the name. Generic matches that start earlier and
          // overlap the name are still blocked by the reservation.
          if (
            !nameMatch ||
            findLongestGenericMatchLength(dictionaryEntries, textWindow) > nameMatch.sourceLength
          ) {
            namePos += String.fromCodePoint(codePoint).length;
            continue;
          }
          const source = text.substring(namePos, namePos + nameMatch.sourceLength);
          nameTokens.push(buildScanToken(namePos, source, {
            term: nameMatch.headword.term,
            reading: nameMatch.headword.reading,
            wordClasses: normalizeWordClasses(nameMatch.headword),
            isNameMatch: true,
            frequencyRank: getBestFrequencyRank(
              nameMatch.dictionaryEntry,
              nameMatch.headwordIndex,
              dictionaryPriorityByName,
              dictionaryFrequencyModeByName
            )
          }));
          namePos += nameMatch.sourceLength;
        }
      }
      let i = 0;
      let nameIndex = 0;
      let unparsedRunStart = null;
      while (i < text.length) {
        while (nameIndex < nameTokens.length && nameTokens[nameIndex].startPos < i) { nameIndex += 1; }
        const nextNameToken = nameIndex < nameTokens.length ? nameTokens[nameIndex] : null;
        if (nextNameToken && nextNameToken.startPos === i) {
          flushUnparsedRun(unparsedRunStart, i);
          unparsedRunStart = null;
          tokens.push(nextNameToken);
          i = nextNameToken.endPos;
          nameIndex += 1;
          continue;
        }
        const codePoint = text.codePointAt(i);
        // Punctuation and whitespace can never start a token: skip the backend
        // round trip entirely. Latin letters and digits stay lookup-worthy
        // (terms like Tシャツ start on an ASCII letter).
        if (!isLookupWorthyCodePoint(codePoint)) {
          if (unparsedRunStart === null) { unparsedRunStart = i; }
          i += String.fromCodePoint(codePoint).length;
          continue;
        }
        // Cap the window at the next reserved name span so a generic match
        // cannot consume into it.
        const windowLength = nextNameToken ? Math.min(scanLength, nextNameToken.startPos - i) : scanLength;
        let attempt = await findTokenAt(i, windowLength);
        // Yomitan text normalization can consume characters (whitespace,
        // punctuation) beyond the matched term, leaving no headword whose
        // source equals the consumed text. Retry with shorter windows so a
        // valid prefix term (e.g. a character name before a paren) still
        // tokenizes instead of the position being skipped. The ladder is
        // capped: without a cap it degrades to O(scanLength) lookups at a
        // single position.
        let retryLength = Math.min(attempt.matchedLength, windowLength) - 1;
        let retryLookupsRemaining = MAX_SHRINKING_WINDOW_RETRY_LOOKUPS;
        while (!attempt.token && retryLength >= 1 && retryLookupsRemaining > 0) {
          retryLookupsRemaining -= 1;
          const retry = await findTokenAt(i, retryLength);
          if (retry.token) {
            attempt = retry;
            break;
          }
          retryLength = Math.min(retryLength - 1, retry.matchedLength - 1);
        }
        if (attempt.token) {
          flushUnparsedRun(unparsedRunStart, i);
          unparsedRunStart = null;
          tokens.push(attempt.token);
          i += attempt.matchedLength;
          continue;
        }
        if (unparsedRunStart === null) { unparsedRunStart = i; }
        i += String.fromCodePoint(text.codePointAt(i)).length;
      }
      flushUnparsedRun(unparsedRunStart, text.length);
      return tokens;
    };
    return true;
  })();
`;

// Installs (or clears) the character-name candidate forms for the current
// media. Runs only when the list changes, not per line. Passing null restores
// the exhaustive every-position pre-pass.
export function buildYomitanScanNameCandidatesScript(
  nameCandidates: { key: string; forms: string[] } | null,
): string {
  if (!nameCandidates) {
    return `
    (() => {
      if (typeof globalThis.__subminerYomitanScanSetNameCandidates !== "function") {
        return false;
      }
      return globalThis.__subminerYomitanScanSetNameCandidates(null, null);
    })();
  `;
  }

  return `
    (() => {
      if (typeof globalThis.__subminerYomitanScanSetNameCandidates !== "function") {
        return false;
      }
      return globalThis.__subminerYomitanScanSetNameCandidates(
        ${JSON.stringify(nameCandidates.key)},
        ${JSON.stringify(nameCandidates.forms)}
      );
    })();
  `;
}

export function buildYomitanScanCallScript(params: YomitanScanRequestParams): string {
  return `
    (async () => {
      if (typeof globalThis.__subminerYomitanScan !== "function") {
        return ${JSON.stringify(YOMITAN_SCAN_RUNTIME_MISSING_SENTINEL)};
      }
      return await globalThis.__subminerYomitanScan(${JSON.stringify(params)});
    })();
  `;
}
