import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  readSubtitleAudioCache,
  writeSubtitleAudioCache,
  subtitleAudioCacheKey,
} from './subtitle-generation-audio-cache';

test('audio cache isolates tracks and headers, verifies content and preserves active copies', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'subtitle-audio-cache-'));
  try {
    const headers = { headers: { Referer: 'https://example.test/' }, userAgent: null };
    const key = subtitleAudioCacheKey('https://example.test/video?secret=token', 2, headers);
    assert.match(key, /^[a-f0-9]{64}$/);
    assert.notEqual(
      key,
      subtitleAudioCacheKey('https://example.test/video?secret=token', 3, headers),
    );
    assert.notEqual(
      key,
      subtitleAudioCacheKey('https://example.test/video?secret=token', 2, {
        ...headers,
        userAgent: 'different',
      }),
    );
    const wavPath = path.join(directory, 'original.wav');
    const destination = path.join(directory, 'active.wav');
    const bytes = Buffer.alloc(1024, 42);
    await writeFile(wavPath, bytes);
    await writeSubtitleAudioCache({ directory, key, wavPath, offset: 0.125 });
    assert.deepEqual(await readSubtitleAudioCache({ directory, key, destination }), {
      offset: 0.125,
    });
    assert.deepEqual(await readFile(destination), bytes);
    const entry = path.join(directory, 'audio', key);
    await writeFile(path.join(entry, 'audio.wav'), Buffer.alloc(1024, 43));
    assert.equal(await readSubtitleAudioCache({ directory, key, destination }), null);
    await assert.rejects(readFile(destination), /ENOENT/);
    await writeSubtitleAudioCache({ directory, key, wavPath, offset: 0.125 });
    assert.ok(await readSubtitleAudioCache({ directory, key, destination }));
    await rm(entry, { recursive: true });
    assert.deepEqual(await readFile(destination), bytes);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
