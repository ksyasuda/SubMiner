import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { createSubtitleGenerationSession } from './subtitle-generation-session';

test('session cleanup removes audio and work files, recovers crashes and preserves results and live sessions', async () => {
  const root = await mkdtemp(path.join(tmpdir(), 'subtitle-session-test-'));
  const exited = spawn(process.execPath, ['-e', ''], { stdio: 'ignore' });
  await new Promise<void>((resolve, reject) => {
    exited.once('error', reject);
    exited.once('close', () => resolve());
  });
  assert.ok(exited.pid);
  const old = path.join(root, 'sessions', `${exited.pid}-abandoned`);
  const live = path.join(root, 'sessions', `${process.pid}-other`);
  try {
    for (const dir of [old, live, path.join(root, 'audio')]) {
      await mkdir(dir, { recursive: true });
      await writeFile(path.join(dir, 'audio.wav'), 'audio');
    }
    const saved = path.join(root, 'episode.ja.generated.srt');
    await writeFile(saved, 'saved subtitles');
    const session = createSubtitleGenerationSession(root);
    const directory = await session.directory();
    assert.equal(await session.directory(), directory);
    await assert.rejects(readFile(path.join(old, 'audio.wav')), /ENOENT/);
    await assert.rejects(readFile(path.join(root, 'audio', 'audio.wav')), /ENOENT/);
    await writeFile(path.join(directory, 'working.wav'), 'temporary');
    await session.dispose();
    await session.dispose();
    await assert.rejects(readdir(directory), /ENOENT/);
    await assert.rejects(session.directory(), /closed/);
    assert.equal(await readFile(saved, 'utf8'), 'saved subtitles');
    assert.equal(await readFile(path.join(live, 'audio.wav'), 'utf8'), 'audio');
    const next = createSubtitleGenerationSession(root);
    assert.notEqual(await next.directory(), directory);
    await next.dispose();
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});
