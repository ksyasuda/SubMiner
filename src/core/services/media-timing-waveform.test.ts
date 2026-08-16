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

test('waveform peaks are normalized without flattening quieter sections', () => {
  const peaks = computeWaveformPeaks(pcm([0, 1_000, -2_000, 4_000, -8_000, 16_000]), 3);

  assert.equal(peaks.length, 3);
  assert.ok((peaks[0] ?? 0) > 0);
  assert.ok((peaks[0] ?? 0) < (peaks[1] ?? 0));
  assert.ok((peaks[1] ?? 0) < (peaks[2] ?? 0));
  assert.equal(peaks[2], 1);
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
