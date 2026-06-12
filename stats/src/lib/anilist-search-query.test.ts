import assert from 'node:assert/strict';
import test from 'node:test';
import { normalizeAnilistSearchQuery } from './anilist-search-query';

test('normalizeAnilistSearchQuery removes appended season scope from anime titles', () => {
  assert.equal(normalizeAnilistSearchQuery('Sword Art Online Season 1'), 'Sword Art Online');
  assert.equal(normalizeAnilistSearchQuery('KonoSuba Season 02'), 'KonoSuba');
});

test('normalizeAnilistSearchQuery removes bracketed season scope without dropping real title text', () => {
  assert.equal(normalizeAnilistSearchQuery('KonoSuba (Season 2)'), 'KonoSuba');
  assert.equal(normalizeAnilistSearchQuery('KonoSuba - Season 2'), 'KonoSuba');
});

test('normalizeAnilistSearchQuery keeps inputs when stripping season scope would erase title', () => {
  assert.equal(normalizeAnilistSearchQuery('Season 1'), 'Season 1');
});
