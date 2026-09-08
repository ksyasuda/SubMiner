import assert from 'node:assert/strict';
import test from 'node:test';
import { parseSpeechPassages, speechPassageCues } from './subtitle-generation-speech';

test('speech passages convert centiseconds, group nearby speech, and retain long gaps', () => {
  assert.deepEqual(
    parseSpeechPassages(
      [
        'Detected 4 speech segments:',
        'Speech segment 0: start = 0.00, end = 300.00',
        'Speech segment 1: start = 350.00, end = 900.00',
        'Speech segment 2: start = 10000.00, end = 11900.00',
        'Speech segment 3: start = 11950.00, end = 12500.00',
      ].join('\n'),
    ),
    [
      { startSeconds: 0, endSeconds: 9 },
      { startSeconds: 100, endSeconds: 119 },
      { startSeconds: 119.5, endSeconds: 125 },
    ],
  );
});

test('speech detector distinguishes no speech from missing, malformed, or truncated output', () => {
  assert.deepEqual(parseSpeechPassages('Detected 0 speech segments:'), []);
  assert.deepEqual(
    parseSpeechPassages(
      'Detected 1 speech segments:\nSpeech segment 0: start = 79896.00, end = 82218.00',
    ),
    [{ startSeconds: 798.96, endSeconds: 822.18 }],
  );
  for (const output of [
    '',
    'Detected 1 speech segments:',
    'Detected 1 speech segments:\nSpeech segment 1: start = 100.00, end = 200.00',
    'Detected 1 speech segments:\nSpeech segment 0: start = 200.00, end = 100.00',
    'Detected 1 speech segments:\nSpeech segment 0: start = NaN, end = 100.00',
    'Detected 1 speech segments:\nSpeech segment 0: start = 0.00, end = 10000.00',
    'Detected 2 speech segments:\nSpeech segment 0: start = 0.00, end = 200.00\nSpeech segment 1: start = 100.00, end = 300.00',
  ])
    assert.throws(() => parseSpeechPassages(output), /Speech detector/);
});

test('passage cue times cannot extend into omitted audio or accumulate offsets', () => {
  const srt =
    '1\n00:00:00,000 --> 00:00:01,000\nはい\n\n2\n00:00:01,000 --> 00:01:40,000\nはい\n\n3\n00:01:41,000 --> 00:01:42,000\n幻覚\n';
  assert.deepEqual(
    speechPassageCues(srt, { startSeconds: 1200.25, endSeconds: 1203.75 }).map(
      ({ startTime, endTime, text }) => ({ startTime, endTime, text }),
    ),
    [
      { startTime: 1200.25, endTime: 1201.25, text: 'はい' },
      { startTime: 1201.25, endTime: 1203.75, text: 'はい' },
    ],
  );
  assert.deepEqual(speechPassageCues('', { startSeconds: 0, endSeconds: 1 }), []);
  assert.throws(
    () => speechPassageCues('broken SRT', { startSeconds: 0, endSeconds: 1 }),
    /malformed/,
  );
});
