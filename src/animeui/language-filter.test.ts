import test from 'node:test';
import assert from 'node:assert/strict';
import {
  collectLanguages,
  filterByLanguage,
  languageLabel,
  pruneSelection,
  toggleLanguage,
} from './language-filter';

const ext = (lang: string, name = lang) => ({ lang, name });

test('languageLabel names known tags and falls back to the raw code', () => {
  assert.equal(languageLabel('ja'), 'Japanese');
  assert.equal(languageLabel('all'), 'Multi-language');
  assert.equal(languageLabel('zzzz'), 'ZZZZ');
});

test('collectLanguages dedupes, puts multi-language first, then sorts by name', () => {
  // English, German, Japanese — display names, not codes.
  assert.deepEqual(collectLanguages([ext('ja'), ext('de'), ext('ja'), ext('all'), ext('en')]), [
    'all',
    'en',
    'de',
    'ja',
  ]);
});

test('toggleLanguage adds, removes, and empties back to All', () => {
  const one = toggleLanguage(new Set(), 'ja');
  assert.deepEqual([...one], ['ja']);
  const two = toggleLanguage(one, 'en');
  assert.deepEqual([...two].sort(), ['en', 'ja']);
  assert.deepEqual([...toggleLanguage(two, 'ja')], ['en']);
  assert.equal(toggleLanguage(one, 'ja').size, 0);
});

test('pruneSelection drops languages no repository offers any more', () => {
  assert.deepEqual([...pruneSelection(new Set(['ja', 'de']), ['ja', 'en'])], ['ja']);
});

test('filterByLanguage keeps everything when nothing is selected', () => {
  const extensions = [ext('ja'), ext('en'), ext('all')];
  assert.deepEqual(filterByLanguage(extensions, new Set()), extensions);
  assert.deepEqual(
    filterByLanguage(extensions, new Set(['ja', 'all'])).map((extension) => extension.name),
    ['ja', 'all'],
  );
});
