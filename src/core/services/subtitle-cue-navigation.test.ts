import assert from 'node:assert/strict';
import test from 'node:test';
import {
  resolveSanitizedSubtitleSeekCommand,
  subtitleCueSeekTime,
} from './subtitle-cue-navigation';

test('next subtitle navigation skips generated ASS events and seeks to the next sanitized cue', () => {
  const cues = [
    {
      startTime: 10,
      endTime: 13,
      text: 'first lyric',
      source: 'canonical-ass' as const,
      animationStartTime: 9.7,
      animationEndTime: 13.4,
    },
    {
      startTime: 13,
      endTime: 16,
      text: 'second lyric',
      source: 'canonical-ass' as const,
      animationStartTime: 12.7,
      animationEndTime: 16.4,
    },
  ];

  assert.deepEqual(resolveSanitizedSubtitleSeekCommand(['sub-seek', 1], cues, 10.2), [
    'seek',
    13.08,
    'absolute+exact',
  ]);
});

test('next subtitle navigation treats simultaneous sanitized cues as one line boundary', () => {
  const cues = [
    { startTime: 10, endTime: 13, text: 'romaji' },
    { startTime: 10.02, endTime: 13, text: 'English' },
    { startTime: 13, endTime: 16, text: 'next romaji' },
    { startTime: 13.02, endTime: 16, text: 'Next English' },
  ];

  assert.deepEqual(resolveSanitizedSubtitleSeekCommand(['sub-seek', 1], cues, 10.1), [
    'seek',
    13.08,
    'absolute+exact',
  ]);
});

test('next subtitle navigation advances past the latest overlapping lyric', () => {
  const cues = [
    { startTime: 10, endTime: 14, text: 'exiting lyric' },
    { startTime: 13, endTime: 16, text: 'current lyric' },
    { startTime: 16, endTime: 19, text: 'next lyric' },
  ];

  assert.deepEqual(resolveSanitizedSubtitleSeekCommand(['sub-seek', 1], cues, 13.2), [
    'seek',
    16.08,
    'absolute+exact',
  ]);
});

test('previous subtitle navigation leaves the current cue and seeks to the prior cue', () => {
  const cues = [
    { startTime: 10, endTime: 12, text: 'first line' },
    { startTime: 13, endTime: 16, text: 'current line' },
  ];

  assert.deepEqual(resolveSanitizedSubtitleSeekCommand(['sub-seek', -1], cues, 14.5), [
    'seek',
    10.08,
    'absolute+exact',
  ]);
});

test('subtitle navigation falls back when no sanitized destination exists', () => {
  const cues = [{ startTime: 10, endTime: 13, text: 'only line' }];

  assert.equal(resolveSanitizedSubtitleSeekCommand(['sub-seek', 1], cues, 10.2), null);
  assert.equal(resolveSanitizedSubtitleSeekCommand(['seek', 5], cues, 10.2), null);
});

test('sidebar cue seeks share the boundary-safe sanitized cue timestamp', () => {
  assert.equal(subtitleCueSeekTime({ startTime: 1, endTime: 2, text: 'line' }), 1.08);
  assert.equal(subtitleCueSeekTime({ startTime: 1, endTime: 1.04, text: 'short' }), 1.03);
});
