import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { findAudiblePassages, mergeSpeechPassages } from './subtitle-generation-coverage';

async function analyze(lines: string[], progress = 'out_time_us=20000000\n') {
  const directory = await mkdtemp(path.join(tmpdir(), 'subtitle-coverage-test-'));
  try {
    const ffmpegPath = path.join(directory, 'ffmpeg');
    await writeFile(
      ffmpegPath,
      `#!${process.execPath}
process.stderr.write(${JSON.stringify(lines.join('\n') + '\n')});
process.stdout.write(${JSON.stringify(progress)});
`,
      { mode: 0o755 },
    );
    return await findAudiblePassages({ ffmpegPath, wavPath: 'audio.wav' });
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

test('audible coverage retains the full timeline when there is no confident silence', async () => {
  assert.deepEqual(await analyze([]), [{ startSeconds: 0, endSeconds: 20 }]);
});

test('audible coverage omits silence while padding nearby audio without exceeding the timeline', async () => {
  assert.deepEqual(
    await analyze([
      '[silencedetect] silence_start: 0',
      '[silencedetect] silence_end: 2 | silence_duration: 2',
      '[silencedetect] silence_start: 8',
      '[silencedetect] silence_end: 12 | silence_duration: 4',
      '[silencedetect] silence_start: 18',
    ]),
    [
      { startSeconds: 1.65, endSeconds: 8.35 },
      { startSeconds: 11.65, endSeconds: 18.35 },
    ],
  );
  assert.deepEqual(await analyze(['[silencedetect] silence_end: 10.5 | silence_duration: 0.5']), [
    { startSeconds: 0, endSeconds: 20 },
  ]);
});

test('entirely silent audio has no audible passages', async () => {
  assert.deepEqual(await analyze(['[silencedetect] silence_start: 0']), []);
  assert.deepEqual(await analyze(['[silencedetect] silence_end: 20 | silence_duration: 20']), []);
});

test('missing analysis duration fails instead of silently dropping audio', async () => {
  await assert.rejects(analyze([], ''), /valid duration/);
});

test('merging coverage preserves quiet VAD speech and does not mutate detector results', () => {
  const speech = [{ startSeconds: 10, endSeconds: 11 }];
  assert.deepEqual(
    mergeSpeechPassages([
      ...speech,
      { startSeconds: 0, endSeconds: 5 },
      { startSeconds: 4, endSeconds: 8 },
      { startSeconds: 11, endSeconds: 12 },
    ]),
    [
      { startSeconds: 0, endSeconds: 8 },
      { startSeconds: 10, endSeconds: 12 },
    ],
  );
  assert.deepEqual(speech, [{ startSeconds: 10, endSeconds: 11 }]);
});
