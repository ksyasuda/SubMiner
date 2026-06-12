import assert from 'node:assert/strict';
import test from 'node:test';
import { buildFfmpegSubtitleExtractionArgs } from './internal-subtitle-extraction';

test('buildFfmpegSubtitleExtractionArgs rejects output paths without an extension', () => {
  assert.throws(
    () => buildFfmpegSubtitleExtractionArgs('/tmp/video.mkv', 2, '/tmp/subtitle-output'),
    /outputPath.*file extension/,
  );
});
