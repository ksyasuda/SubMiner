import test from 'node:test';
import assert from 'node:assert/strict';
import { classifyChapterSkipIntervals } from './aniskip-chapters';

const DURATION = 1440;

function chapters(...entries: Array<[number, string]>) {
  return entries.map(([time, title]) => ({ time, title }));
}

test('finds named openings and endings case-insensitively', () => {
  const result = classifyChapterSkipIntervals(
    chapters([0, 'Prologue'], [95, 'OPENING'], [185, 'Part A'], [1310, 'Ending'], [1400, 'Next']),
    DURATION,
  );
  assert.deepEqual(result, {
    intro: { start: 95, end: 185 },
    outro: { start: 1310, end: 1400 },
  });
});

test('matches short OP/ED titles but not words that merely start with them', () => {
  assert.deepEqual(
    classifyChapterSkipIntervals(chapters([0, 'Op'], [90, 'Episode'], [1350, 'ED']), DURATION),
    { intro: { start: 0, end: 90 }, outro: { start: 1350, end: 1440 } },
  );
  assert.deepEqual(
    classifyChapterSkipIntervals(
      chapters([0, 'OP "Song Title"'], [90, 'Episode'], [1350, 'ED "Song Title"']),
      DURATION,
    ),
    { intro: { start: 0, end: 90 }, outro: { start: 1350, end: 1440 } },
  );
  assert.deepEqual(
    classifyChapterSkipIntervals(chapters([0, 'Operation'], [500, 'Editing']), DURATION),
    { intro: null, outro: null },
  );
});

test('recognises Japanese chapter names', () => {
  assert.deepEqual(
    classifyChapterSkipIntervals(
      chapters([0, 'アバン'], [60, 'オープニング'], [150, '本編'], [1330, 'エンディング']),
      DURATION,
    ),
    { intro: { start: 60, end: 150 }, outro: { start: 1330, end: 1440 } },
  );
});

test('accepts outro and credits names, with bare "Credits" only in the second half', () => {
  assert.deepEqual(
    classifyChapterSkipIntervals(chapters([0, 'Episode'], [1350, 'Outro']), DURATION).outro,
    { start: 1350, end: 1440 },
  );
  assert.deepEqual(
    classifyChapterSkipIntervals(chapters([0, 'Episode'], [1350, 'Closing Credits']), DURATION)
      .outro,
    { start: 1350, end: 1440 },
  );
  assert.deepEqual(
    classifyChapterSkipIntervals(chapters([0, 'Credits'], [90, 'Episode']), DURATION).outro,
    null,
  );
  assert.deepEqual(
    classifyChapterSkipIntervals(chapters([0, 'Episode'], [1350, 'Credits']), DURATION).outro,
    { start: 1350, end: 1440 },
  );
});

test('falls back to song-length heuristics when the file has at least three chapters', () => {
  const result = classifyChapterSkipIntervals(
    chapters(
      [0, 'Chapter 1'],
      [60, 'Chapter 2'], // 60 s long, in the first 30%
      [120, 'Chapter 3'], // 88 s long, nearest 90 s
      [208, 'Chapter 4'],
      [1100, 'Chapter 5'], // starts after 70% but 230 s long
      [1330, 'Chapter 6'], // first song-length chapter after 70%
      [1420, 'Chapter 7'],
    ),
    DURATION,
  );
  assert.deepEqual(result, {
    intro: { start: 120, end: 208 },
    outro: { start: 1330, end: 1420 },
  });
});

test('skips heuristics for files with fewer than three chapters', () => {
  assert.deepEqual(
    classifyChapterSkipIntervals(chapters([0, 'Chapter 1'], [1350, 'Chapter 2']), DURATION),
    { intro: null, outro: null },
  );
});

test("ignores SubMiner's own AniSkip marker chapters", () => {
  const result = classifyChapterSkipIntervals(
    chapters(
      [0, 'Episode'],
      [10, 'AniSkip Intro Start'],
      [100, 'AniSkip Intro End'],
      [1330, 'AniSkip Outro Start'],
      [1420, 'AniSkip Outro End'],
    ),
    DURATION,
  );
  assert.deepEqual(result, { intro: null, outro: null });
});

test('returns nothing without a usable duration or chapter list', () => {
  assert.deepEqual(classifyChapterSkipIntervals(chapters([0, 'Opening'], [90, 'A']), 0), {
    intro: null,
    outro: null,
  });
  assert.deepEqual(classifyChapterSkipIntervals(null, DURATION), { intro: null, outro: null });
});
