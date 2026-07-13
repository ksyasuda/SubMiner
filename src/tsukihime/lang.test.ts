import assert from 'node:assert/strict';
import test from 'node:test';

import {
  tsukihimeLangToFilenameSuffix,
  tsukihimeTrackMatchesLanguages,
  describeTsukihimeTabLanguages,
  normalizeTsukihimeLangCode,
} from './lang.js';

test('normalizeTsukihimeLangCode collapses 2/3-letter and region variants', () => {
  assert.equal(normalizeTsukihimeLangCode('eng'), 'en');
  assert.equal(normalizeTsukihimeLangCode('en'), 'en');
  assert.equal(normalizeTsukihimeLangCode('en-US'), 'en');
  assert.equal(normalizeTsukihimeLangCode('GER'), 'de');
  assert.equal(normalizeTsukihimeLangCode('jpn'), 'ja');
  assert.equal(normalizeTsukihimeLangCode('vie'), 'vie');
  assert.equal(normalizeTsukihimeLangCode(''), '');
});

test('tsukihimeTrackMatchesLanguages matches across code forms', () => {
  assert.equal(tsukihimeTrackMatchesLanguages('eng', ['en']), true);
  assert.equal(tsukihimeTrackMatchesLanguages('eng', ['en', 'eng']), true);
  assert.equal(tsukihimeTrackMatchesLanguages('ger', ['de']), true);
  assert.equal(tsukihimeTrackMatchesLanguages('por', ['en']), false);
  assert.equal(tsukihimeTrackMatchesLanguages('spa', ['en', 'de']), false);
});

test('tsukihimeTrackMatchesLanguages keeps unknown-language tracks visible', () => {
  assert.equal(tsukihimeTrackMatchesLanguages('', ['en']), true);
  assert.equal(tsukihimeTrackMatchesLanguages('und', ['en']), true);
});

test('describeTsukihimeTabLanguages names common languages and dedupes', () => {
  assert.equal(describeTsukihimeTabLanguages(['en', 'eng']), 'English');
  assert.equal(describeTsukihimeTabLanguages(['de']), 'German');
  assert.equal(describeTsukihimeTabLanguages(['en', 'de']), 'English / German');
  assert.equal(describeTsukihimeTabLanguages(['vie']), 'VIE');
  assert.equal(describeTsukihimeTabLanguages([]), 'English');
});

test('tsukihimeLangToFilenameSuffix is re-exported from the pure module', () => {
  assert.equal(tsukihimeLangToFilenameSuffix('jpn'), 'ja');
  assert.equal(tsukihimeLangToFilenameSuffix('eng'), 'en');
});
