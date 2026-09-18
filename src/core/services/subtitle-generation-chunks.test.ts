import assert from 'node:assert/strict';
import test from 'node:test';
import { appendSpeechChunkCues, splitSpeechPassages } from './subtitle-generation-chunks';

test('reference starts guide long cuts ahead of VAD without dropping unreferenced audio', () => {
  const chunks = splitSpeechPassages(
    [{ startSeconds: 0, endSeconds: 70 }],
    [18],
    [19],
    [22, 43, 200],
  );
  assert.deepEqual(chunks, [
    { startSeconds: 0, endSeconds: 22.25 },
    { startSeconds: 21.75, endSeconds: 43.25 },
    { startSeconds: 42.75, endSeconds: 63.25 },
    { startSeconds: 62.75, endSeconds: 70 },
  ]);
  assert.deepEqual(splitSpeechPassages([{ startSeconds: 0, endSeconds: 25 }], [], [], [10, 20]), [
    { startSeconds: 0, endSeconds: 25 },
  ]);
});

test('long coverage cuts at nearby speech starts instead of leaving a quiet lead-in', () => {
  const chunks = splitSpeechPassages(
    [{ startSeconds: 544.418, endSeconds: 581.581 }],
    [563.928],
    [555.07, 567.23, 579.91],
  );
  assert.deepEqual(chunks, [
    { startSeconds: 544.418, endSeconds: 567.48 },
    { startSeconds: 566.98, endSeconds: 581.581 },
  ]);
});

test('short audible passages stay intact even with several detected speech starts', () => {
  assert.deepEqual(
    splitSpeechPassages(
      [{ startSeconds: 876.897, endSeconds: 897.812 }],
      [],
      [876.9, 881.15, 893.95, 897.99],
    ),
    [{ startSeconds: 876.897, endSeconds: 897.812 }],
  );
});

test('speech anchors outside retained coverage cannot extend a chunk across a silent gap', () => {
  const chunks = splitSpeechPassages(
    [{ startSeconds: 100, endSeconds: 142 }],
    [118],
    [90, 142, 144],
  );
  assert.deepEqual(chunks, [
    { startSeconds: 100, endSeconds: 118.25 },
    { startSeconds: 117.75, endSeconds: 138.25 },
    { startSeconds: 137.75, endSeconds: 142 },
  ]);
});

test('long speech splits near a pause with context on both sides and no lost audio', () => {
  assert.deepEqual(splitSpeechPassages([{ startSeconds: 100, endSeconds: 145 }], [105, 118, 137]), [
    { startSeconds: 100, endSeconds: 118.25 },
    { startSeconds: 117.75, endSeconds: 137.25 },
    { startSeconds: 136.75, endSeconds: 145 },
  ]);
});

test('uninterrupted speech retains overlapping context without crossing omitted gaps', () => {
  const chunks = splitSpeechPassages([
    { startSeconds: 0, endSeconds: 60 },
    { startSeconds: 100, endSeconds: 100.15 },
  ]);
  assert.deepEqual(chunks, [
    { startSeconds: 0, endSeconds: 20.25 },
    { startSeconds: 19.75, endSeconds: 40.25 },
    { startSeconds: 39.75, endSeconds: 60 },
    { startSeconds: 100, endSeconds: 100.15 },
  ]);
});

test('speech fitting one Whisper window stays intact instead of cutting a sentence at 20 seconds', () => {
  const passage = { startSeconds: 251.71, endSeconds: 275.01 };
  assert.deepEqual(splitSpeechPassages([passage], [271.327]), [passage]);
});

test('chunk stitching ignores punctuation differences without merging separate repetitions', () => {
  const cues = [{ startTime: 19.7, endTime: 21.2, text: 'ありがとう' }];
  appendSpeechChunkCues(cues, [
    { startTime: 19.8, endTime: 21.3, text: 'ありがとう。' },
    { startTime: 22, endTime: 23, text: 'ありがとう！' },
  ]);
  assert.deepEqual(cues, [
    { startTime: 19.7, endTime: 21.3, text: 'ありがとう' },
    { startTime: 22, endTime: 23, text: 'ありがとう！' },
  ]);
});

test('chunk stitching removes matching overlap cues but retains repeated dialogue', () => {
  const cues = [{ startTime: 19.7, endTime: 20.2, text: 'はい' }];
  appendSpeechChunkCues(cues, [
    { startTime: 19.8, endTime: 20.3, text: 'はい' },
    { startTime: 21, endTime: 21.5, text: 'はい' },
    { startTime: 21.4, endTime: 22, text: 'はい' },
  ]);
  assert.deepEqual(cues, [
    { startTime: 19.7, endTime: 20.3, text: 'はい' },
    { startTime: 21, endTime: 21.5, text: 'はい' },
    { startTime: 21.4, endTime: 22, text: 'はい' },
  ]);
});

test('chunk stitching matches repeated text to the greatest overlap without leaving a duplicate', () => {
  const cues = [
    { startTime: 10, endTime: 14, text: 'はい' },
    { startTime: 13, endTime: 20, text: 'はい' },
  ];
  appendSpeechChunkCues(cues, [
    { startTime: 12, endTime: 21, text: 'はい' },
    { startTime: 22, endTime: 23, text: 'はい' },
  ]);
  assert.deepEqual(cues, [
    { startTime: 10, endTime: 14, text: 'はい' },
    { startTime: 12, endTime: 21, text: 'はい' },
    { startTime: 22, endTime: 23, text: 'はい' },
  ]);
});
