import assert from 'node:assert/strict';
import test from 'node:test';
import { appendSpeechChunkCues, splitSpeechPassages } from './subtitle-generation-chunks';

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
