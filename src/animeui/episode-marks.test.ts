import test from 'node:test';
import assert from 'node:assert/strict';
import { describeMarkCount, episodesInScope } from './episode-marks';

const LIST = ['e12', 'e11', 'e10', 'e9'];

test('the one scope takes only the episode itself', () => {
  assert.deepEqual(episodesInScope(LIST, 1, 'one'), ['e11']);
});

test('the below scope takes the episode and everything listed after it', () => {
  assert.deepEqual(episodesInScope(LIST, 1, 'below'), ['e11', 'e10', 'e9']);
  assert.deepEqual(episodesInScope(LIST, 0, 'below'), LIST);
  assert.deepEqual(episodesInScope(LIST, 3, 'below'), ['e9']);
});

test('an index outside the list marks nothing', () => {
  assert.deepEqual(episodesInScope(LIST, -1, 'below'), []);
  assert.deepEqual(episodesInScope(LIST, 4, 'one'), []);
});

test('describeMarkCount agrees with itself about plurals and direction', () => {
  assert.equal(describeMarkCount(1, true), 'Marked 1 episode watched');
  assert.equal(describeMarkCount(3, true), 'Marked 3 episodes watched');
  assert.equal(describeMarkCount(1, false), 'Cleared the watch mark on 1 episode');
  assert.equal(describeMarkCount(12, false), 'Cleared the watch mark on 12 episodes');
});
