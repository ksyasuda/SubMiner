import test from 'node:test';
import assert from 'node:assert/strict';

import { buildWhisperArgs } from './whisper';

test('buildWhisperArgs includes threads and optional VAD flags', () => {
  assert.deepEqual(
    buildWhisperArgs({
      modelPath: '/models/ggml-large-v2.bin',
      audioPath: '/tmp/input.wav',
      outputPrefix: '/tmp/output',
      language: 'ja',
      translate: false,
      threads: 8,
      vadModelPath: '/models/vad.bin',
    }),
    [
      '-m',
      '/models/ggml-large-v2.bin',
      '-f',
      '/tmp/input.wav',
      '--output-srt',
      '--output-file',
      '/tmp/output',
      '--language',
      'ja',
      '--threads',
      '8',
      '-vm',
      '/models/vad.bin',
      '--vad',
    ],
  );
});

test('buildWhisperArgs includes translate flag when requested', () => {
  assert.ok(
    buildWhisperArgs({
      modelPath: '/models/base.bin',
      audioPath: '/tmp/input.wav',
      outputPrefix: '/tmp/output',
      language: 'ja',
      translate: true,
      threads: 4,
    }).includes('--translate'),
  );
});
