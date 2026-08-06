// Match selection for the injected scan runtime: picks the headword a position
// tokenizes to, and the longest name or generic match in a window, which is how
// the greedy name pre-pass decides what to reserve.

export const YOMITAN_MATCH_SELECTION_HELPERS = String.raw`
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
          // Every match already comes from currentMediaDictionaryEntries, so
          // classifying its own entry is enough.
          for (const match of exactPrimaryMatches) {
            if (!isCurrentMediaNameDictionaryEntry(match.dictionaryEntry)) { continue; }
            matchedNameDictionary = true;
            break;
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
