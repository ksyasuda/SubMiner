import assert from 'node:assert/strict';
import test from 'node:test';

import {
  clampCharacterCursor,
  countCharacters,
  moveCharacterCursor,
} from './keyboard-character-cursor.js';

// もっとハ | キ | ハキ: the split that hides ハキハキ from word navigation.
const LENGTHS = [4, 1, 2];

test('the character cursor walks a word and crosses into the next one', () => {
  assert.deepEqual(moveCharacterCursor(LENGTHS, { wordIndex: 0, charIndex: 2 }, 1), {
    position: { wordIndex: 0, charIndex: 3 },
    result: 'moved',
  });
  assert.deepEqual(moveCharacterCursor(LENGTHS, { wordIndex: 0, charIndex: 3 }, 1), {
    position: { wordIndex: 1, charIndex: 0 },
    result: 'moved',
  });
  assert.deepEqual(moveCharacterCursor(LENGTHS, { wordIndex: 2, charIndex: 0 }, -1), {
    position: { wordIndex: 1, charIndex: 0 },
    result: 'moved',
  });
  assert.deepEqual(moveCharacterCursor(LENGTHS, { wordIndex: 1, charIndex: 0 }, -1), {
    position: { wordIndex: 0, charIndex: 3 },
    result: 'moved',
  });
});

test('the character cursor reports the line edges and stays put there', () => {
  assert.deepEqual(moveCharacterCursor(LENGTHS, { wordIndex: 0, charIndex: 0 }, -1), {
    position: { wordIndex: 0, charIndex: 0 },
    result: 'start-boundary',
  });
  assert.deepEqual(moveCharacterCursor(LENGTHS, { wordIndex: 2, charIndex: 1 }, 1), {
    position: { wordIndex: 2, charIndex: 1 },
    result: 'end-boundary',
  });
  assert.equal(moveCharacterCursor([], { wordIndex: 0, charIndex: 0 }, 1).result, 'no-words');
});

test('words without characters are skipped', () => {
  assert.deepEqual(moveCharacterCursor([1, 0, 2], { wordIndex: 0, charIndex: 0 }, 1).position, {
    wordIndex: 2,
    charIndex: 0,
  });
  assert.deepEqual(moveCharacterCursor([1, 0, 2], { wordIndex: 2, charIndex: 0 }, -1).position, {
    wordIndex: 0,
    charIndex: 0,
  });
});

test('a position past the end lands on the last character', () => {
  assert.deepEqual(
    clampCharacterCursor(LENGTHS, { wordIndex: 9, charIndex: Number.MAX_SAFE_INTEGER }),
    { wordIndex: 2, charIndex: 1 },
  );
  assert.equal(clampCharacterCursor([], { wordIndex: 0, charIndex: 0 }), null);
});

test('characters are counted as code points', () => {
  assert.equal(countCharacters('𠮷野家'), 3);
  assert.equal(countCharacters('ハキハキ'), 4);
});
