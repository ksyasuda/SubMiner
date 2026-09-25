import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { detectSubtitleGenerationAcceleration } from './subtitle-generation-acceleration';

async function fixture(run: (directory: string) => Promise<void>) {
  const directory = await mkdtemp(path.join(tmpdir(), 'subtitle-acceleration-test-'));
  try {
    await run(directory);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

async function executable(directory: string, name: string, body: string) {
  const file = path.join(directory, name);
  await writeFile(file, `#!${process.execPath}\n${body}`, { mode: 0o755 });
  return file;
}

const cudaOutput =
  'ggml_cuda_init: found 1 CUDA devices:\nwhisper_model_load: invalid model data (bad magic)\n';

test('NVIDIA and CUDA discovery work without downloading or loading a model', () =>
  fixture(async (directory) => {
    await executable(directory, 'nvidia-smi', 'console.log("NVIDIA Test GPU");');
    const whisper = await executable(
      directory,
      'whisper-cli',
      `const fs = require('node:fs');
       const model = process.argv[process.argv.indexOf('-m') + 1];
       if (fs.readFileSync(model).length !== 4) process.exit(1);
       fs.writeFileSync(${JSON.stringify(path.join(directory, 'probe-path'))}, model);
       process.stderr.write(${JSON.stringify(cudaOutput)}); process.exit(3);`,
    );
    assert.deepEqual(
      await detectSubtitleGenerationAcceleration(
        { kind: 'found', path: whisper },
        { PATH: directory },
      ),
      { kind: 'nvidia-cuda', gpuName: 'NVIDIA Test GPU' },
    );
    const model = await readFile(path.join(directory, 'probe-path'), 'utf8');
    await assert.rejects(readdir(path.dirname(model)), { code: 'ENOENT' });
  }));

test('CPU-only, Vulkan-only, hidden CUDA devices and incomplete probes fall back safely', () =>
  fixture(async (directory) => {
    await executable(directory, 'nvidia-smi', 'console.log("NVIDIA Test GPU");');
    for (const output of [
      'usage: --no-gpu disable GPU\ninvalid model data (bad magic)',
      'ggml_vulkan: Found 1 Vulkan devices\ninvalid model data (bad magic)',
      'ggml_cuda_init: found 0 CUDA devices\ninvalid model data (bad magic)',
      'ggml_cuda_init: found 1 CUDA devices\nCUDA error: driver initialization failed',
    ]) {
      const whisper = await executable(
        directory,
        'whisper-cli',
        `process.stderr.write(${JSON.stringify(output)}); process.exit(3);`,
      );
      assert.deepEqual(
        await detectSubtitleGenerationAcceleration(
          { kind: 'found', path: whisper },
          { PATH: directory },
        ),
        { kind: 'unavailable' },
      );
    }
  }));

test('missing tools, missing NVIDIA devices and driver errors do not recommend turbo', () =>
  fixture(async (directory) => {
    const whisper = await executable(
      directory,
      'whisper-cli',
      `process.stderr.write(${JSON.stringify(cudaOutput)}); process.exit(3);`,
    );
    assert.deepEqual(
      await detectSubtitleGenerationAcceleration(
        { kind: 'missing', message: 'Not installed' },
        { PATH: directory },
      ),
      { kind: 'unavailable' },
    );
    for (const driver of [
      null,
      'process.exit(0);',
      'console.log("NVIDIA GPU"); process.exit(1);',
    ]) {
      if (driver !== null) await executable(directory, 'nvidia-smi', driver);
      assert.deepEqual(
        await detectSubtitleGenerationAcceleration(
          { kind: 'found', path: whisper },
          { PATH: directory },
        ),
        { kind: 'unavailable' },
      );
    }
  }));

test('a hung Whisper probe times out even if it printed a CUDA device', () =>
  fixture(async (directory) => {
    await executable(directory, 'nvidia-smi', 'console.log("NVIDIA Test GPU");');
    const whisper = await executable(
      directory,
      'whisper-cli',
      `process.stderr.write(${JSON.stringify(cudaOutput)}); setInterval(() => {}, 1000);`,
    );
    const started = Date.now();
    assert.deepEqual(
      await detectSubtitleGenerationAcceleration(
        { kind: 'found', path: whisper },
        { PATH: directory },
      ),
      { kind: 'unavailable' },
    );
    assert.ok(Date.now() - started < 6000);
  }));
