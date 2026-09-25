import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import test from 'node:test';
import {
  buildFfmpegSubtitleExtractionArgs,
  createCachedInternalSubtitleTrackExtractor,
  extractInternalSubtitleTrackToTempFile,
  parseTrackId,
} from './internal-subtitle-extraction';

test('buildFfmpegSubtitleExtractionArgs rejects output paths without an extension', () => {
  assert.throws(
    () => buildFfmpegSubtitleExtractionArgs('/tmp/video.mkv', 2, '/tmp/subtitle-output'),
    /outputPath.*file extension/,
  );
});

test('parseTrackId rejects negative track ids', () => {
  assert.equal(parseTrackId(-1), null);
  assert.equal(parseTrackId(' -2 '), null);
});

test('cached internal subtitle extraction shares concurrent and repeated track requests', async () => {
  let extractionCalls = 0;
  let cleanupCalls = 0;
  let resolveExtraction:
    | ((result: { path: string; cleanup: () => Promise<void> }) => void)
    | undefined;
  const firstExtraction = new Promise<{ path: string; cleanup: () => Promise<void> }>((resolve) => {
    resolveExtraction = resolve;
  });
  const extractor = createCachedInternalSubtitleTrackExtractor({
    extract: async () => {
      extractionCalls += 1;
      if (extractionCalls === 1) {
        return firstExtraction;
      }
      return {
        path: `/tmp/subtitle-${extractionCalls}.ass`,
        cleanup: async () => {
          cleanupCalls += 1;
        },
      };
    },
  });
  const request = () =>
    extractor.extract('ffmpeg', '/Volumes/media/episode.mkv', {
      'ff-index': 3,
      codec: 'ass',
    });

  const concurrent = Array.from({ length: 6 }, request);
  assert.equal(extractionCalls, 1);
  if (!resolveExtraction) {
    throw new Error('extraction did not start');
  }
  resolveExtraction({
    path: '/tmp/subtitle-1.ass',
    cleanup: async () => {
      cleanupCalls += 1;
    },
  });

  const results = await Promise.all(concurrent);
  assert.deepEqual(
    results.map((result) => result?.path),
    Array.from({ length: 6 }, () => '/tmp/subtitle-1.ass'),
  );
  await Promise.all(results.map((result) => result?.cleanup()));
  assert.equal(cleanupCalls, 0);

  assert.equal((await request())?.path, '/tmp/subtitle-1.ass');
  assert.equal(extractionCalls, 1);

  extractor.clear();
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(cleanupCalls, 1);
  assert.equal((await request())?.path, '/tmp/subtitle-2.ass');
  assert.equal(extractionCalls, 2);
});

test('extractInternalSubtitleTrackToTempFile times out stalled ffmpeg process', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-ffmpeg-timeout-'));
  const videoPath = path.join(root, 'video.mkv');
  fs.writeFileSync(videoPath, '');

  try {
    await assert.rejects(
      () =>
        extractInternalSubtitleTrackToTempFile(
          process.execPath,
          videoPath,
          { 'ff-index': 0, codec: 'ass' },
          {
            extractionTimeoutMs: 20,
            spawnArgsOverride: ['-e', 'setTimeout(() => {}, 1000);'],
          },
        ),
      /ffmpeg extraction timed out/,
    );
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});
