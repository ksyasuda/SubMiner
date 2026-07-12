import assert from 'node:assert/strict';
import test from 'node:test';

import {
  animetoshoLangToFilenameSuffix,
  animetoshoTrackMatchesLanguages,
  describeAnimetoshoTabLanguages,
  normalizeAnimetoshoLangCode,
} from './lang.js';

test('normalizeAnimetoshoLangCode collapses 2/3-letter and region variants', () => {
  assert.equal(normalizeAnimetoshoLangCode('eng'), 'en');
  assert.equal(normalizeAnimetoshoLangCode('en'), 'en');
  assert.equal(normalizeAnimetoshoLangCode('en-US'), 'en');
  assert.equal(normalizeAnimetoshoLangCode('GER'), 'de');
  assert.equal(normalizeAnimetoshoLangCode('jpn'), 'ja');
  assert.equal(normalizeAnimetoshoLangCode('vie'), 'vie');
  assert.equal(normalizeAnimetoshoLangCode(''), '');
});

test('animetoshoTrackMatchesLanguages matches across code forms', () => {
  assert.equal(animetoshoTrackMatchesLanguages('eng', ['en']), true);
  assert.equal(animetoshoTrackMatchesLanguages('eng', ['en', 'eng']), true);
  assert.equal(animetoshoTrackMatchesLanguages('ger', ['de']), true);
  assert.equal(animetoshoTrackMatchesLanguages('por', ['en']), false);
  assert.equal(animetoshoTrackMatchesLanguages('spa', ['en', 'de']), false);
});

test('animetoshoTrackMatchesLanguages keeps unknown-language tracks visible', () => {
  assert.equal(animetoshoTrackMatchesLanguages('', ['en']), true);
  assert.equal(animetoshoTrackMatchesLanguages('und', ['en']), true);
});

test('describeAnimetoshoTabLanguages names common languages and dedupes', () => {
  assert.equal(describeAnimetoshoTabLanguages(['en', 'eng']), 'English');
  assert.equal(describeAnimetoshoTabLanguages(['de']), 'German');
  assert.equal(describeAnimetoshoTabLanguages(['en', 'de']), 'English / German');
  assert.equal(describeAnimetoshoTabLanguages(['vie']), 'VIE');
  assert.equal(describeAnimetoshoTabLanguages([]), 'English');
});

test('animetoshoLangToFilenameSuffix is re-exported from the pure module', () => {
  assert.equal(animetoshoLangToFilenameSuffix('jpn'), 'ja');
  assert.equal(animetoshoLangToFilenameSuffix('eng'), 'en');
});
