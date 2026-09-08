import assert from 'node:assert/strict';
import test from 'node:test';
import {
  describeGenerationModel,
  describeGenerationProgress,
  describeGenerationVad,
} from './subtitle-generation-view';

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
