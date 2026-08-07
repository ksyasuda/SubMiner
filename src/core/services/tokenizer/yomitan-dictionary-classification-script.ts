// Dictionary classification for the injected scan runtime: which dictionaries
// an entry came from, and whether it is a SubMiner character entry for the
// media being watched. Both walk nested entry data, so both are memoized on the
// entry object by the runtime that hosts them.

import { CHARACTER_DICTIONARY_TITLE_PREFIX } from './character-dictionary-title';

// The prefix is interpolated into generated regex source, so metacharacters in
// it would change what the pattern matches (or fail to compile).
const ESCAPED_TITLE_PREFIX = CHARACTER_DICTIONARY_TITLE_PREFIX.replace(
  /[.*+?^${}()|[\]\\]/g,
  '\\$&',
);
const TITLE_MEDIA_ID_PATTERN = ESCAPED_TITLE_PREFIX + String.raw`[^\d]*(?:AniList\s*)?(\d+)`;

export const YOMITAN_DICTIONARY_CLASSIFICATION_HELPERS = String.raw`
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
        const TITLE_MEDIA_ID_REGEX = new RegExp(${JSON.stringify(TITLE_MEDIA_ID_PATTERN)}, 'i');
        function parseSubMinerMediaIdFromString(value) {
          const imageMatch = value.match(/\bimg\/m(\d+)-/i);
          if (imageMatch) {
            const parsed = Number.parseInt(imageMatch[1], 10);
            if (Number.isSafeInteger(parsed) && parsed > 0) { return parsed; }
          }
          const titleMatch = value.match(TITLE_MEDIA_ID_REGEX);
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
`;
