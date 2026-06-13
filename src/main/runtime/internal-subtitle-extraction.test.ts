import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import test from 'node:test';
import {
  buildFfmpegSubtitleExtractionArgs,
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
