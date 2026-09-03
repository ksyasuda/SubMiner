import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildSpeechWaveformArgs,
  computeWaveformPeaks,
  generateSpeechWaveform,
} from './media-timing-waveform';

function pcm(samples: number[]): Buffer {
  const result = Buffer.alloc(samples.length * 2);
  samples.forEach((sample, index) => result.writeInt16LE(sample, index * 2));
  return result;
}

test('speech waveform maps the selected FFmpeg stream and visible range', () => {
  const args = buildSpeechWaveformArgs(
    {
      mediaPath: '/video/show.mkv',
      startTime: 8,
      endTime: 15,
      audioStreamIndex: 3,
    },
    'center',
  );

  assert.deepEqual(args.slice(args.indexOf('-ss'), args.indexOf('-t') + 2), [
    '-ss',
    '8',
    '-i',
    '/video/show.mkv',
    '-t',
    '7',
  ]);
  assert.deepEqual(args.slice(args.indexOf('-map'), args.indexOf('-map') + 2), ['-map', '0:3']);
  assert.match(args[args.indexOf('-af') + 1] ?? '', /c0=FC/);
});

test('speech waveform seeks cached windows by source timestamps', () => {
  const args = buildSpeechWaveformArgs(
    {
      mediaPath: { path: '/tmp/window.mkv', absoluteTimestamps: true, singleResolvedStream: true },
      startTime: 8,
      endTime: 15,
    },
    'downmix',
  );

  assert.deepEqual(args.slice(args.indexOf('-ss'), args.indexOf('-t') + 2), [
    '-ss',
    '8',
    '-seek_timestamp',
    '1',
    '-i',
    '/tmp/window.mkv',
    '-t',
    '7',
  ]);
  assert.equal(args.includes('-map'), false);
});

test('waveform levels rise with loudness and top out at the reference level', () => {
  const peaks = computeWaveformPeaks(pcm([0, 1_000, -2_000, 4_000, -8_000, 16_000]), 3);

  assert.equal(peaks.length, 3);
  assert.equal(peaks[0], 0);
  assert.ok((peaks[1] ?? 0) > 0);
  assert.ok((peaks[1] ?? 0) < (peaks[2] ?? 0));
  assert.equal(peaks[2], 1);
});

test('waveform flattens steady background noise and keeps speech bursts tall', () => {
  // 20 slices of steady noise at a fixed level with an 18 dB louder "speech" burst in the middle.
  const noise = 1_000;
  const samples: number[] = [];
  for (let slice = 0; slice < 20; slice += 1) {
    const level = slice >= 8 && slice < 12 ? noise * 8 : noise;
    for (let sample = 0; sample < 50; sample += 1) {
      samples.push(sample % 2 === 0 ? level : -level);
    }
  }

  const peaks = computeWaveformPeaks(pcm(samples), 20);

  for (const [index, peak] of peaks.entries()) {
    if (index >= 8 && index < 12) assert.equal(peak, 1);
    else assert.equal(peak, 0);
  }
});

test('waveform stays flat when the whole range is a single steady level', () => {
  const peaks = computeWaveformPeaks(
    pcm(Array.from({ length: 400 }, (_, i) => (i % 2 ? 900 : -900))),
    40,
  );

  assert.ok(peaks.every((peak) => peak === 0));
});

test('speech waveform uses a mono downmix when the source has no center activity', async () => {
  const calls: string[][] = [];
  const peaks = await generateSpeechWaveform(
    { mediaPath: '/video/show.mkv', startTime: 0, endTime: 2 },
    async (args) => {
      calls.push(args);
      return calls.length === 1 ? pcm([0, 0, 0, 0]) : pcm([0, 4_000, -8_000, 16_000]);
    },
  );

  assert.equal(calls.length, 2);
  assert.match(calls[1]?.[calls[1].indexOf('-af') + 1] ?? '', /channel_layouts=mono/);
  assert.equal(Math.max(...peaks), 1);
});

test('speech waveform keeps an active center channel without doing a second decode', async () => {
  let calls = 0;
  await generateSpeechWaveform(
    { mediaPath: '/video/show.mkv', startTime: 0, endTime: 2 },
    async () => {
      calls += 1;
      return pcm([0, 4_000, -8_000, 16_000]);
    },
  );

  assert.equal(calls, 1);
});
