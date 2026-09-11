import assert from 'node:assert/strict';
import test from 'node:test';
import {
  describeGenerationModel,
  describeGenerationProgress,
  describeGenerationTools,
  describeGenerationVad,
} from './subtitle-generation-view';

test('missing tools block generation and list every install instruction', () => {
  const found = { kind: 'found', path: '/usr/bin/tool' } as const;
  assert.deepEqual(
    describeGenerationTools({ ffmpeg: found, ffprobe: found, whisper: found, vad: null }),
    {
      ready: true,
      text: 'whisper.cpp and FFmpeg are installed.',
    },
  );
  assert.deepEqual(
    describeGenerationTools({
      ffmpeg: { kind: 'missing', message: 'ffmpeg was not found on PATH.' },
      ffprobe: found,
      whisper: found,
      vad: { kind: 'missing', message: 'whisper-vad-speech-segments was not found on PATH.' },
    }),
    {
      ready: false,
      text: 'ffmpeg was not found on PATH. whisper-vad-speech-segments was not found on PATH.',
    },
  );
});

test('only a missing managed model offers a download', () => {
  assert.deepEqual(describeGenerationModel({ kind: 'missing', path: '/models/small.bin' }), {
    ready: false,
    download: true,
    text: 'Download a speech model to get started.',
  });
  for (const kind of ['managed', 'external'] as const) {
    const model = describeGenerationModel({ kind, path: '/models/ggml-small.bin' });
    assert.equal(model.ready, true);
    assert.equal(model.download, false);
  }
  assert.deepEqual(
    describeGenerationModel({
      kind: 'invalid',
      path: '/missing/model.bin',
      message: 'Configured model does not exist.',
    }),
    {
      ready: false,
      download: false,
      text: 'Configured model does not exist.',
    },
  );
});

test('generation progress distinguishes measured work from indeterminate stages', () => {
  assert.deepEqual(
    describeGenerationProgress({ stage: 'transcribe', percent: 42.8, message: 'Transcribing' }),
    {
      stage: 'Recognizing Japanese speech',
      percent: 42.8,
      label: '42%',
    },
  );
  assert.deepEqual(describeGenerationProgress({ stage: 'extract', message: 'Extracting audio' }), {
    stage: 'Preparing audio',
    percent: null,
    label: 'Working...',
  });
  assert.equal(
    describeGenerationProgress({ stage: 'download', percent: 0, message: '' }).label,
    '0%',
  );
  assert.equal(
    describeGenerationProgress({ stage: 'download', percent: Number.NaN, message: '' }).percent,
    null,
  );
  assert.equal(
    describeGenerationProgress({ stage: 'write', percent: 120, message: '' }).percent,
    100,
  );
});

test('optional speech detection only gates generation when selected', () => {
  const missing = { kind: 'missing', path: '/vad.bin' } as const;
  assert.equal(describeGenerationVad({ enabled: false, model: missing }).ready, true);
  assert.equal(describeGenerationVad({ enabled: false, model: missing }).download, false);
  assert.equal(describeGenerationVad({ enabled: true, model: missing }).ready, false);
  assert.equal(describeGenerationVad({ enabled: true, model: missing }).download, true);
  assert.equal(
    describeGenerationVad({ enabled: true, model: { kind: 'managed', path: '/vad.bin' } }).ready,
    true,
  );
  const invalid = { kind: 'invalid', path: '/vad.bin', message: 'Cannot read model' } as const;
  assert.equal(describeGenerationVad({ enabled: true, model: invalid }).download, false);
  assert.equal(describeGenerationVad({ enabled: false, model: invalid }).ready, true);
});
