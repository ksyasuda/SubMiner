import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { downloadSubtitleGenerationArtifact } from './subtitle-generation-download';
import {
  downloadSubtitleGenerationVadModel,
  resolveSubtitleGenerationVadModel,
} from './subtitle-generation-vad-model';
import { DEFAULT_SUBTITLE_GENERATION_CONFIG } from '../../shared/subtitle-generation';

test('verified model publication preserves existing files and cleans up failed or cancelled downloads', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'subtitle-model-download-'));
  const originalFetch = globalThis.fetch;
  const bytes = new TextEncoder().encode('fixture model');
  const input = {
    url: 'https://example.test/model',
    size: bytes.length,
    sha256: createHash('sha256').update(bytes).digest('hex'),
    destination: path.join(directory, 'model.bin'),
    label: 'test model',
  };
  try {
    globalThis.fetch = Object.assign(async () => new Response(bytes), originalFetch);
    await downloadSubtitleGenerationArtifact(input);
    assert.equal(await readFile(input.destination, 'utf8'), 'fixture model');
    await assert.rejects(downloadSubtitleGenerationArtifact(input), /EEXIST/);
    await assert.rejects(
      downloadSubtitleGenerationArtifact({
        ...input,
        destination: path.join(directory, 'bad.bin'),
        sha256: 'wrong',
      }),
      /integrity/,
    );
    const controller = new AbortController();
    await assert.rejects(
      downloadSubtitleGenerationArtifact({
        ...input,
        destination: path.join(directory, 'cancelled.bin'),
        signal: controller.signal,
        onProgress: ({ percent }) => {
          if (percent === 99) controller.abort();
        },
      }),
    );
    assert.deepEqual(await readdir(directory), ['model.bin']);
  } finally {
    globalThis.fetch = originalFetch;
    await rm(directory, { recursive: true, force: true });
  }
});

test('VAD setup recognizes existing paths and never replaces an invalid external model', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'subtitle-vad-model-'));
  try {
    const config = { ...DEFAULT_SUBTITLE_GENERATION_CONFIG };
    assert.equal((await resolveSubtitleGenerationVadModel(config, directory)).kind, 'missing');
    config.vadModelPath = path.join(directory, 'external.bin');
    await assert.rejects(
      downloadSubtitleGenerationVadModel({ config, modelDirectory: directory }),
      /Cannot read/,
    );
    await writeFile(config.vadModelPath, 'external model');
    assert.equal((await resolveSubtitleGenerationVadModel(config, directory)).kind, 'external');
    assert.equal(
      await downloadSubtitleGenerationVadModel({ config, modelDirectory: directory }),
      config.vadModelPath,
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
