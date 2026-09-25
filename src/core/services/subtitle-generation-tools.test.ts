import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { DEFAULT_SUBTITLE_GENERATION_CONFIG } from '../../shared/subtitle-generation';
import {
  requireSubtitleGenerationTools,
  resolveSubtitleGenerationTools,
} from './subtitle-generation-tools';

async function fixture(run: (directory: string) => Promise<void>) {
  const directory = await mkdtemp(path.join(tmpdir(), 'subtitle-generation-tools-'));
  try {
    await run(directory);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

async function executable(directory: string, name: string): Promise<string> {
  const file = path.join(directory, name);
  await writeFile(file, '#!/bin/sh\n', { mode: 0o755 });
  return file;
}

test('tools resolve from PATH, honor overrides, and only require the detector in dialogue mode', () =>
  fixture(async (directory) => {
    const bin = path.join(directory, 'bin');
    await mkdir(bin);
    for (const name of ['ffmpeg', 'ffprobe', 'whisper-cli']) await executable(bin, name);
    const detector = await executable(bin, 'vad-speech-segments');
    const customWhisper = await executable(directory, 'my-whisper');
    await writeFile(path.join(directory, 'not-executable'), '', { mode: 0o644 });
    const env = { PATH: bin };

    const found = await resolveSubtitleGenerationTools(DEFAULT_SUBTITLE_GENERATION_CONFIG, env);
    assert.deepEqual(found, {
      ffmpeg: { kind: 'found', path: path.join(bin, 'ffmpeg') },
      ffprobe: { kind: 'found', path: path.join(bin, 'ffprobe') },
      whisper: { kind: 'found', path: path.join(bin, 'whisper-cli') },
      vad: null,
    });
    assert.equal(requireSubtitleGenerationTools(found).vad, null);

    const dialogue = await resolveSubtitleGenerationTools(
      { ...DEFAULT_SUBTITLE_GENERATION_CONFIG, vadModelPath: '/models/vad.bin' },
      env,
    );
    assert.deepEqual(dialogue.vad, { kind: 'found', path: detector });

    const overridden = await resolveSubtitleGenerationTools(
      {
        ...DEFAULT_SUBTITLE_GENERATION_CONFIG,
        whisperPath: customWhisper,
        ffmpegPath: path.join(directory, 'not-executable'),
      },
      env,
    );
    assert.deepEqual(overridden.whisper, { kind: 'found', path: customWhisper });
    assert.equal(overridden.ffmpeg.kind, 'missing');
    assert.throws(
      () => requireSubtitleGenerationTools(overridden),
      /not-executable \(subtitleGeneration\.ffmpegPath\) is not an executable file/,
    );
  }));

test('missing tools name the executable, the installer, and the setting', () =>
  fixture(async (directory) => {
    const tools = await resolveSubtitleGenerationTools(
      { ...DEFAULT_SUBTITLE_GENERATION_CONFIG, vadModelPath: '/models/vad.bin' },
      { PATH: directory },
    );
    assert.deepEqual(tools.whisper, {
      kind: 'missing',
      message:
        'whisper-cli was not found on PATH. Install whisper.cpp or set subtitleGeneration.whisperPath in Settings.',
    });
    assert.deepEqual(tools.vad, {
      kind: 'missing',
      message:
        "whisper-vad-speech-segments was not found on PATH. Install whisper.cpp's speech segment detector or set subtitleGeneration.vadPath in Settings.",
    });
    assert.throws(() => requireSubtitleGenerationTools(tools), /ffmpeg was not found on PATH/);
  }));
