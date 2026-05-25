import assert from 'node:assert/strict';
import test from 'node:test';
import { estimateSubtitleTimingOffset } from './subtitle-timing-offset';

function cue(startTime: number) {
  return { startTime, endTime: startTime + 1, text: `cue ${startTime}` };
}

test('estimate subtitle timing offset detects a late Jellyfin subtitle timeline', () => {
  const primary = [
    34.935, 36.937, 41.441, 45.279, 48.115, 52.286, 54.955, 59.793, 63.63, 67.634, 76.643, 80.814,
    87.988, 90.991, 94.094, 97.097,
  ].map(cue);
  const reference = [
    3.46, 9.48, 13.61, 21.4, 28.16, 32.06, 35.93, 45.1, 56.57, 59.68, 62.44, 65.56,
  ].map(cue);

  const result = estimateSubtitleTimingOffset(primary, reference);

  assert.ok(result);
  assert.ok(result.offsetSeconds > -32);
  assert.ok(result.offsetSeconds < -31);
  assert.ok(result.matchCount >= 8);
  assert.ok(result.meanErrorSeconds <= 0.75);
});

test('estimate subtitle timing offset favors the early episode timeline', () => {
  const primary = [
    34.935, 36.937, 41.441, 45.279, 48.115, 52.286, 54.955, 59.793, 63.63, 67.634, 76.643, 80.814,
    87.988, 90.991, 94.094, 97.097, 207.974, 212.579, 222.422, 228.095, 232.432, 238.271, 244.778,
    246.78, 249.282, 251.284, 253.62, 256.289, 259.626, 262.129, 264.965, 267.634, 270.303, 274.407,
    277.077, 280.08, 284.084, 288.421, 291.925, 295.262, 298.431, 301.101, 306.773, 308.942,
    312.946, 316.283, 321.621, 326.626, 331.131, 336.069, 340.407, 343.41, 351.418, 355.422,
    357.924, 362.429, 365.432, 370.604, 373.273, 377.944, 381.114, 384.618, 387.621, 390.957,
    396.73, 399.232, 401.568, 403.57, 405.572, 407.574, 409.743, 412.746, 418.752, 425.258, 427.26,
    435.602, 440.44, 442.942, 445.445, 449.783,
  ].map(cue);
  const reference = [
    3.46, 9.48, 13.61, 21.4, 28.16, 32.06, 35.93, 45.1, 56.57, 59.68, 62.44, 65.56, 165.77, 172.81,
    176.1, 177.27, 186.33, 191.33, 195.78, 201.83, 212.9, 214.09, 216.73, 220.2, 222.91, 225.65,
    232.8, 237.92, 242.23, 243.28, 247.53, 252.04, 255.9, 258.86, 262.09, 264.43, 276.07, 278.01,
    280.98, 285.67, 289.89, 294.57, 300, 303.56, 308.58, 316.37, 318.38, 319.86, 325.38, 328.82,
    333.68, 335.26, 336.82, 340.11, 342.11, 344.36, 346.39, 347.53, 350.92, 370.18, 372.88, 376.43,
    388.2, 390.57, 403.96, 406.36, 409.72, 413.78, 425.55, 432.76, 435.03, 438.06, 443.73, 448.31,
    450.57, 457.62, 463.41, 465.85, 473.79, 480.59,
  ].map(cue);

  const result = estimateSubtitleTimingOffset(primary, reference);

  assert.ok(result);
  assert.ok(result.offsetSeconds > -32);
  assert.ok(result.offsetSeconds < -31);
});

test('estimate subtitle timing offset ignores subtitle timelines that are already aligned', () => {
  const starts = [1, 5, 9, 14, 20, 25, 31, 38];

  const result = estimateSubtitleTimingOffset(
    starts.map(cue),
    starts.map((start) => cue(start + 0.04)),
  );

  assert.equal(result, null);
});

test('estimate subtitle timing offset rejects weak timeline matches', () => {
  const primary = [10, 20, 30, 40, 50, 60, 70, 80].map(cue);
  const reference = [1, 2, 3, 4, 5, 6, 7, 8].map(cue);

  const result = estimateSubtitleTimingOffset(primary, reference);

  assert.equal(result, null);
});
