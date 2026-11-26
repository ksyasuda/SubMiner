import test from 'node:test';
import assert from 'node:assert/strict';
import { detectSpeechSpansFromPcm, parseRawSileroVadSegments } from './audio-speech';

test('detectSpeechSpansFromPcm detects voice-band bursts separated by quiet background', () => {
  const sampleRate = 8000;
  const totalSec = 20;
  const samples = new Int16Array(sampleRate * totalSec);

  const expectedSpans: Array<[number, number]> = [
    [2.0, 4.0],
    [7.5, 9.5],
    [13.0, 15.5],
  ];

  // Fill quiet ambient floor across whole clip.
  for (let i = 0; i < samples.length; i += 1) {
    const t = i / sampleRate;
    samples[i] = Math.round(Math.sin(2 * Math.PI * 500 * t) * 180);
  }

  // Overlay speech-like multi-frequency formant signal inside expected spans.
  for (const [start, end] of expectedSpans) {
    const sIdx = Math.floor(start * sampleRate);
    const eIdx = Math.floor(end * sampleRate);
    for (let i = sIdx; i < eIdx; i += 1) {
      const t = i / sampleRate;
      const sig =
        0.5 * Math.sin(2 * Math.PI * 420 * t) +
        0.3 * Math.sin(2 * Math.PI * 1150 * t) +
        0.2 * Math.sin(2 * Math.PI * 2100 * t);
      samples[i] = Math.round(sig * 12000);
    }
  }

  const { spans, durationSec } = detectSpeechSpansFromPcm(samples, sampleRate);
  assert.equal(durationSec, totalSec);
  assert.equal(spans.length, expectedSpans.length);
  for (let i = 0; i < expectedSpans.length; i += 1) {
    assert.ok(Math.abs(spans[i]!.start - expectedSpans[i]![0]) <= 0.15);
    assert.ok(Math.abs(spans[i]!.end - expectedSpans[i]![1]) <= 0.25);
  }
});

test('parseRawSileroVadSegments parses fine-grained speech segments from whisper-vad output', () => {
  const stdout = [
    'Detected 2 speech segments:',
    'Speech segment 0: start = 125.00, end = 340.00',
    'Speech segment 1: start = 510.50, end = 690.00',
  ].join('\n');

  assert.deepEqual(parseRawSileroVadSegments(stdout), [
    { start: 1.25, end: 3.4 },
    { start: 5.105, end: 6.9 },
  ]);
});
