// Furigana distribution for the injected scan runtime: splits a headword and
// its reading into the segments a token carries, including the inflected case
// where the matched source text differs from the dictionary form.

export const YOMITAN_FURIGANA_HELPERS = String.raw`
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
`;
