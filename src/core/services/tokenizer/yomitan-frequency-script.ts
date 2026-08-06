// Frequency-rank resolution for the injected scan runtime: reads the many
// shapes a Yomitan frequency entry can take and picks the best rank for a
// headword, honouring per-dictionary priority and occurrence-vs-rank mode.

export const YOMITAN_FREQUENCY_HELPERS = String.raw`
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
`;
