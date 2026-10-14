import assert from 'node:assert/strict';
import test from 'node:test';
import type { SubtitleCue } from '../../types';
import { buildOverlapPrefetchCues } from './subtitle-prefetch-overlaps';

test('overlapping cues yield the combined line the live path resolves', () => {
  const cues: SubtitleCue[] = [
    { startTime: 0, endTime: 4, text: 'このまま頑張ったって―' },
    { startTime: 2, endTime: 6, text: 'ちょっとカズマ\n聞こえてんの？' },
  ];

  assert.deepEqual(buildOverlapPrefetchCues(cues), [
    { startTime: 2, endTime: 4, text: 'このまま頑張ったって―\n\nちょっとカズマ\n聞こえてんの？' },
  ]);
});

test('back-to-back cues yield no overlap lines', () => {
  const cues: SubtitleCue[] = [
    { startTime: 0, endTime: 2, text: 'first' },
    { startTime: 2, endTime: 4, text: 'second' },
  ];

  assert.deepEqual(buildOverlapPrefetchCues(cues), []);
});

test('each overlap span joins only the cues active across it, in cue-list order', () => {
  const cues: SubtitleCue[] = [
    { startTime: 3, endTime: 8, text: 'C' },
    { startTime: 0, endTime: 6, text: 'A' },
    { startTime: 2, endTime: 4, text: 'B' },
  ];

  assert.deepEqual(buildOverlapPrefetchCues(cues), [
    { startTime: 2, endTime: 3, text: 'A\n\nB' },
    { startTime: 3, endTime: 4, text: 'C\n\nA\n\nB' },
    { startTime: 4, endTime: 6, text: 'C\n\nA' },
  ]);
});
