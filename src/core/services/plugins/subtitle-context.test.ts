import assert from 'node:assert/strict';
import test from 'node:test';

import { buildPluginSubtitleContext } from './subtitle-context';

const cues = [
  { text: '一', startTime: 1, endTime: 2 },
  { text: 'はい', startTime: 3, endTime: 4 },
  { text: '二', startTime: 5, endTime: 6 },
  { text: '三', startTime: 7, endTime: 8 },
  { text: 'はい', startTime: 9, endTime: 10 },
  { text: '四', startTime: 11, endTime: 12 },
];

test('returns neighbors of the matching cue, preferring the one nearest the live timing', () => {
  // Live timing is shifted by a +1.5 s sub-delay, so it sits outside both はい cues.
  const context = buildPluginSubtitleContext({
    cues,
    current: { text: 'はい', startTime: 10.5, endTime: 11.5 },
    recentLines: [],
    before: 2,
    after: 1,
  });

  assert.deepEqual(
    context.before.map((line) => line.text),
    ['二', '三'],
  );
  assert.deepEqual(
    context.after.map((line) => line.text),
    ['四'],
  );
});

test('falls back to recently displayed lines when the line is not in the cue list', () => {
  const context = buildPluginSubtitleContext({
    cues: [],
    current: { text: '今', startTime: 20, endTime: 21 },
    recentLines: [
      { text: '前の前', startTime: 16, endTime: 17 },
      { text: '前', startTime: 18, endTime: 19 },
      { text: '今', startTime: 20, endTime: 21 },
    ],
    before: 2,
  });

  assert.deepEqual(
    context.before.map((line) => line.text),
    ['前の前', '前'],
  );
  assert.deepEqual(context.after, []);
});

test('returns an empty context when no subtitle is displayed', () => {
  assert.deepEqual(
    buildPluginSubtitleContext({
      cues,
      current: { text: '  ', startTime: null, endTime: null },
      recentLines: [],
    }),
    { current: null, before: [], after: [] },
  );
});
