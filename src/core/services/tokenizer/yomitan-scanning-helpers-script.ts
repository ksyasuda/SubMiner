// Helper bundle for the in-page Yomitan scan runtime: kana/furigana handling,
// headword preference, and frequency-rank resolution. Injected as text into the
// parser window by yomitan-scan-runtime-script.ts, so it is data here, not code
// this process runs.
import { HAN_CODE_POINT_RANGES } from '../../text/han-code-points';

export const CHARACTER_DICTIONARY_TITLE_PREFIX = 'SubMiner Character Dictionary';

export const YOMITAN_SCANNING_HELPERS = String.raw`
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
      function createFuriganaSegment(text, reading) { return {text, reading}; }
      function getSegmentReadingContribution(segment) {
        if (typeof segment.reading === "string" && segment.reading.length > 0) { return segment.reading; }
        const segmentText = typeof segment.text === "string" ? segment.text : "";
        const isKanaOnly = segmentText.length > 0 && [...segmentText].every((char) => isCodePointKana(char.codePointAt(0)));
        return isKanaOnly ? convertHalfwidthKanaToKatakana(segmentText) : "";
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
            case HALFWIDTH_KANA_PROLONGED_SOUND_MARK_CODE_POINT:
              char = "ー";
              if (!keepProlongedSoundMarks && result.length > 0) {
                const char2 = getProlongedHiragana(result[result.length - 1]);
                if (char2 !== null) { char = char2; }
              }
              break;
            default:
              if (isCodePointInRange(codePoint, KATAKANA_CONVERSION_RANGE)) {
                char = String.fromCodePoint(codePoint + offset);
                break;
              }
              // Halfwidth katakana folds too, or a name written that way would
              // match neither a candidate form nor its own reading.
              const halfwidthHiragana = convertHalfwidthKanaCodePointToHiragana(codePoint);
              if (halfwidthHiragana !== null) { char = halfwidthHiragana; }
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
        // Memoized on the entry object: termsFind results are cached across
        // lines, so the same entries come back for every repeated lookup, and
        // each one is classified several times per scan (name pre-pass,
        // headword preference, every retry window).
        function getDictionaryEntryNames(entry) {
          if (!entry || typeof entry !== 'object') { return []; }
          const cached = dictionaryEntryNamesCache.get(entry);
          if (cached !== undefined) { return cached; }
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
          dictionaryEntryNamesCache.set(entry, names);
          return names;
        }
        // Cached per scan rather than per runtime: the answer depends on
        // includeNameMatchMetadata, which is a per-call parameter.
        const nameDictionaryEntryCache = new WeakMap();
        function isNameDictionaryEntry(entry) {
          if (!includeNameMatchMetadata || !entry || typeof entry !== 'object') {
            return false;
          }
          const cached = nameDictionaryEntryCache.get(entry);
          if (cached !== undefined) { return cached; }
          const isName = getDictionaryEntryNames(entry).some((name) => name.startsWith(${JSON.stringify(CHARACTER_DICTIONARY_TITLE_PREFIX)}));
          nameDictionaryEntryCache.set(entry, isName);
          return isName;
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
        // Walking an entry collects media ids from every nested value, so this
        // is the most expensive classification step; memoized on the entry for
        // the same reason as the dictionary names above.
        function getSubMinerMediaIds(entry) {
          if (!entry || typeof entry !== 'object') { return EMPTY_MEDIA_ID_SET; }
          const cached = subMinerMediaIdsCache.get(entry);
          if (cached !== undefined) { return cached; }
          const mediaIds = new Set();
          collectSubMinerMediaIds(entry, mediaIds);
          subMinerMediaIdsCache.set(entry, mediaIds);
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
