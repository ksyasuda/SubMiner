import assert from 'node:assert/strict';
import test from 'node:test';
import { selectAutoplayStartupCue } from './autoplay-subtitle-primer';

test('selectAutoplayStartupCue returns the active cue at the current time', () => {
  assert.deepEqual(
    selectAutoplayStartupCue(
      [
        { startTime: 1, endTime: 3, text: 'first' },
        { startTime: 4, endTime: 5, text: 'second' },
      ],
      2,
      1,
    ),
    { startTime: 1, endTime: 3, text: 'first' },
  );
});

test('selectAutoplayStartupCue returns the next imminent cue before playback starts', () => {
  assert.deepEqual(
    selectAutoplayStartupCue(
      [
        { startTime: 1.2, endTime: 3, text: 'first' },
        { startTime: 4, endTime: 5, text: 'second' },
      ],
      0,
      2,
    ),
    { startTime: 1.2, endTime: 3, text: 'first' },
  );
});

test('selectAutoplayStartupCue does not reveal far future subtitle text', () => {
  assert.equal(
    selectAutoplayStartupCue([{ startTime: 12, endTime: 15, text: 'later' }], 0, 2),
    null,
  );
});

test('selectAutoplayStartupCue skips blank cues', () => {
  assert.deepEqual(
    selectAutoplayStartupCue(
      [
        { startTime: 0, endTime: 1, text: '   ' },
        { startTime: 0.5, endTime: 2, text: 'visible' },
      ],
      0.75,
      1,
    ),
    { startTime: 0.5, endTime: 2, text: 'visible' },
  );
});
