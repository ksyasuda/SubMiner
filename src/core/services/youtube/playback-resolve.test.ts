import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { resolveYoutubePlaybackUrl } from './playback-resolve';

async function withTempDir<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-youtube-playback-resolve-'));
  try {
    return await fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function makeFakeYtDlpScript(dir: string, payload: string): void {
  const scriptPath = path.join(dir, 'yt-dlp');
  const script = `#!/usr/bin/env node
process.stdout.write(${JSON.stringify(payload)});
`;
  fs.writeFileSync(scriptPath, script, 'utf8');
  if (process.platform !== 'win32') {
    fs.chmodSync(scriptPath, 0o755);
  }
  fs.writeFileSync(scriptPath + '.cmd', `@echo off\r\nnode "${scriptPath}"\r\n`, 'utf8');
}

async function withFakeYtDlp<T>(payload: string, fn: () => Promise<T>): Promise<T> {
  return await withTempDir(async (root) => {
    const binDir = path.join(root, 'bin');
    fs.mkdirSync(binDir, { recursive: true });
    makeFakeYtDlpScript(binDir, payload);
    const fakeCommandPath =
      process.platform === 'win32' ? path.join(binDir, 'yt-dlp.cmd') : path.join(binDir, 'yt-dlp');
    const originalCommand = process.env.SUBMINER_YTDLP_BIN;
    process.env.SUBMINER_YTDLP_BIN = fakeCommandPath;
    try {
      return await fn();
    } finally {
      if (originalCommand === undefined) {
        delete process.env.SUBMINER_YTDLP_BIN;
      } else {
        process.env.SUBMINER_YTDLP_BIN = originalCommand;
      }
    }
  });
}

test('resolveYoutubePlaybackUrl returns the first playable URL line', async () => {
  await withFakeYtDlp(
    '\nhttps://manifest.googlevideo.com/api/manifest/hls_playlist/test\nhttps://ignored.example/video\n',
    async () => {
      const result = await resolveYoutubePlaybackUrl('https://www.youtube.com/watch?v=abc123');
      assert.equal(result, 'https://manifest.googlevideo.com/api/manifest/hls_playlist/test');
    },
  );
});

test('resolveYoutubePlaybackUrl rejects when yt-dlp returns no URL', async () => {
  await withFakeYtDlp('\n', async () => {
    await assert.rejects(
      resolveYoutubePlaybackUrl('https://www.youtube.com/watch?v=abc123'),
      /returned empty output/,
    );
  });
});
