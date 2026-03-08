import test from 'node:test';
import assert from 'node:assert/strict';

import { parseSrt, stringifySrt } from './srt';

test('parseSrt reads cue numbering timing and text', () => {
  const cues = parseSrt(`1
00:00:01,000 --> 00:00:02,000
こんにちは

2
00:00:02,500 --> 00:00:03,000
世界
`);

  assert.equal(cues.length, 2);
  assert.equal(cues[0]?.start, '00:00:01,000');
  assert.equal(cues[0]?.end, '00:00:02,000');
  assert.equal(cues[0]?.text, 'こんにちは');
  assert.equal(cues[1]?.text, '世界');
});

test('stringifySrt preserves parseable cue structure', () => {
  const roundTrip = stringifySrt(
    parseSrt(`1
00:00:01,000 --> 00:00:02,000
こんにちは
`),
  );

  assert.match(roundTrip, /1\n00:00:01,000 --> 00:00:02,000\nこんにちは/);
});
