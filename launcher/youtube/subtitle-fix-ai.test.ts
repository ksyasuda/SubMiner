import test from 'node:test';
import assert from 'node:assert/strict';

import { applyFixedCueBatch, parseAiSubtitleFixResponse } from './subtitle-fix-ai';
import { parseSrt } from './srt';

test('applyFixedCueBatch accepts content-only fixes with identical timing', () => {
  const original = parseSrt(`1
00:00:01,000 --> 00:00:02,000
こんいちは

2
00:00:03,000 --> 00:00:04,000
世界
`);
  const fixed = parseSrt(`1
00:00:01,000 --> 00:00:02,000
こんにちは

2
00:00:03,000 --> 00:00:04,000
世界
`);

  const merged = applyFixedCueBatch(original, fixed);
  assert.equal(merged[0]?.text, 'こんにちは');
});

test('applyFixedCueBatch rejects changed timestamps', () => {
  const original = parseSrt(`1
00:00:01,000 --> 00:00:02,000
こんいちは
`);
  const fixed = parseSrt(`1
00:00:01,100 --> 00:00:02,000
こんにちは
`);

  assert.throws(() => applyFixedCueBatch(original, fixed), /timestamps/i);
});

test('parseAiSubtitleFixResponse accepts valid SRT wrapped in markdown fences', () => {
  const original = parseSrt(`1
00:00:01,000 --> 00:00:02,000
こんいちは

2
00:00:03,000 --> 00:00:04,000
世界
`);

  const parsed = parseAiSubtitleFixResponse(
    original,
    '```srt\n1\n00:00:01,000 --> 00:00:02,000\nこんにちは\n\n2\n00:00:03,000 --> 00:00:04,000\n世界\n```',
  );

  assert.equal(parsed[0]?.text, 'こんにちは');
  assert.equal(parsed[1]?.text, '世界');
});

test('parseAiSubtitleFixResponse accepts text-only one-block-per-cue output', () => {
  const original = parseSrt(`1
00:00:01,000 --> 00:00:02,000
こんいちは

2
00:00:03,000 --> 00:00:04,000
世界
`);

  const parsed = parseAiSubtitleFixResponse(
    original,
    `こんにちは

世界`,
  );

  assert.equal(parsed[0]?.start, '00:00:01,000');
  assert.equal(parsed[0]?.text, 'こんにちは');
  assert.equal(parsed[1]?.end, '00:00:04,000');
  assert.equal(parsed[1]?.text, '世界');
});

test('parseAiSubtitleFixResponse rejects unrecoverable text-only output', () => {
  const original = parseSrt(`1
00:00:01,000 --> 00:00:02,000
こんいちは

2
00:00:03,000 --> 00:00:04,000
世界
`);

  assert.throws(
    () => parseAiSubtitleFixResponse(original, 'こんにちは\n世界\n余分です'),
    /cue block|cue count/i,
  );
});

test('parseAiSubtitleFixResponse rejects language drift for primary Japanese subtitles', () => {
  const original = parseSrt(`1
00:00:01,000 --> 00:00:02,000
こんにちは

2
00:00:03,000 --> 00:00:04,000
今日はいい天気ですね
`);

  assert.throws(
    () =>
      parseAiSubtitleFixResponse(
        original,
        `1
00:00:01,000 --> 00:00:02,000
Hello

2
00:00:03,000 --> 00:00:04,000
The weather is nice today
`,
        'ja',
      ),
    /language/i,
  );
});
