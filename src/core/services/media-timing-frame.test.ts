import assert from 'node:assert/strict';
import test from 'node:test';
import { createMediaTimingFrameExtractor, selectMediaTimingFrame } from './media-timing-frame';

test('frame stepping follows decoded timestamps with variable frame durations', () => {
  const times = [10, 10.041667, 10.125, 10.166667];
  assert.equal(selectMediaTimingFrame(times, 10.05), 10.125);
  assert.equal(selectMediaTimingFrame(times, 10.125 - 0.000001, 1), 10.166667);
  assert.equal(selectMediaTimingFrame(times, 10.125 - 0.000001, -1), 10.041667);
  assert.equal(selectMediaTimingFrame(times, 10, -1), undefined);
  assert.equal(selectMediaTimingFrame(times, 10.166667, 1), undefined);
  assert.equal(selectMediaTimingFrame([], 10), undefined);
});

function fixture(startTime = 0) {
  const calls: Array<{ file: string; args: string[] }> = [];
  const extractor = createMediaTimingFrameExtractor(async (file, args) => {
    calls.push({ file, args });
    if (file === 'ffmpeg') return Buffer.from('image');
    if (args.includes('format=start_time'))
      return Buffer.from(JSON.stringify({ format: { start_time: String(startTime) } }));
    return Buffer.from(
      JSON.stringify({
        frames: [10, 10.04, 10.12, 10.16].map((time) => ({
          best_effort_timestamp_time: String(time + startTime),
        })),
      }),
    );
  });
  return { calls, extractor };
}

test('frame extraction normalizes nonzero source start times and reuses the frame index', async () => {
  const { calls, extractor } = fixture(5);
  const first = await extractor.generate({ media: '/movie.mkv', timestamp: 10.05 });
  assert.ok(Math.abs(first.timestamp - 10.119999) < 0.000001);
  assert.equal(first.dataUrl, 'data:image/jpeg;base64,aW1hZ2U=');
  await extractor.generate({ media: '/movie.mkv', timestamp: first.timestamp, direction: 1 });
  assert.equal(calls.filter((call) => call.file === 'ffprobe').length, 2);
  assert.ok(calls[1]!.args.includes('13.05%17.05'));
  extractor.clear();
  await extractor.generate({ media: '/movie.mkv', timestamp: 10.05 });
  assert.equal(calls.filter((call) => call.file === 'ffprobe').length, 4);
});

test('cached windows keep source timestamps for ffmpeg and use absolute ffprobe intervals', async () => {
  const { calls, extractor } = fixture();
  await extractor.generate({
    media: { path: '/window.mkv', absoluteTimestamps: true },
    timestamp: 10.05,
  });
  assert.ok(calls[1]!.args.includes('8.05%12.05'));
  assert.equal(calls[1]!.args.includes('-seek_timestamp'), false);
  assert.ok(calls[2]!.args.includes('-seek_timestamp'));
});

test('remote frame reads carry source headers and do not silently reuse a different input', async () => {
  const { calls, extractor } = fixture();
  const media = {
    path: 'https://example.test/video',
    inputOptions: { headers: { 'X-Emby-Token': 'test-token' } },
  };
  await extractor.generate({ media, timestamp: 10.05 });
  assert.ok(calls.every((call) => call.args.includes('X-Emby-Token: test-token\r\n')));
  await extractor.generate({
    media: { ...media, path: 'https://example.test/other' },
    timestamp: 10.05,
  });
  assert.equal(calls.filter((call) => call.file === 'ffprobe').length, 4);
});
