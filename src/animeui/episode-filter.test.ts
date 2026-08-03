import test from 'node:test';
import assert from 'node:assert/strict';
import { filterEpisodes, parseEpisodeFilter } from './episode-filter';

const episode = (number: number | null, name: string) => ({ number, name });

const SEASON = [
  episode(1, 'Episode 1 - Departure'),
  episode(2, 'Episode 2 - The Long Road'),
  episode(12, 'Episode 12 - Homecoming'),
  episode(12.5, 'Episode 12.5 - Recap'),
  episode(18, 'Episode 18 - Finale'),
  episode(null, 'OVA: Beach Special'),
];

test('an empty query keeps the whole list', () => {
  assert.deepEqual(filterEpisodes(SEASON, '   '), SEASON);
  assert.equal(parseEpisodeFilter(''), null);
});

test('a number matches that episode, not every episode containing the digits', () => {
  assert.deepEqual(
    filterEpisodes(SEASON, '12').map((item) => item.number),
    [12, 12.5],
  );
});

test('a number still matches names when the source reported no numbers', () => {
  const unnumbered = [episode(null, 'Episode 3'), episode(null, 'Episode 4')];
  assert.deepEqual(
    filterEpisodes(unnumbered, '3').map((item) => item.name),
    ['Episode 3'],
  );
});

test('a range keeps the episodes inside it, in either order', () => {
  assert.deepEqual(
    filterEpisodes(SEASON, '12-18').map((item) => item.number),
    [12, 12.5, 18],
  );
  assert.deepEqual(
    filterEpisodes(SEASON, '18 – 12').map((item) => item.number),
    [12, 12.5, 18],
  );
});

test('a range skips episodes the source gave no number', () => {
  assert.deepEqual(filterEpisodes([episode(null, 'OVA')], '1-99'), []);
});

test('anything else is a case-insensitive substring of the name', () => {
  assert.deepEqual(
    filterEpisodes(SEASON, 'home').map((item) => item.number),
    [12],
  );
  assert.deepEqual(
    filterEpisodes(SEASON, 'BEACH').map((item) => item.name),
    ['OVA: Beach Special'],
  );
  assert.deepEqual(filterEpisodes(SEASON, 'nothing here'), []);
});
