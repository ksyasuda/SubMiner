import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { probeYoutubeTracks } from './track-probe';

async function withTempDir<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-youtube-track-probe-'));
  try {
    return await fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function makeFakeYtDlpScript(dir: string, payload: unknown): void {
  const scriptPath = path.join(dir, 'yt-dlp');
  const script = `#!/usr/bin/env node
process.stdout.write(${JSON.stringify(JSON.stringify(payload))});
`;
  fs.writeFileSync(scriptPath, script, 'utf8');
  fs.chmodSync(scriptPath, 0o755);
}

async function withFakeYtDlp<T>(payload: unknown, fn: () => Promise<T>): Promise<T> {
  return await withTempDir(async (root) => {
    const binDir = path.join(root, 'bin');
    fs.mkdirSync(binDir, { recursive: true });
    makeFakeYtDlpScript(binDir, payload);
    const originalPath = process.env.PATH ?? '';
    process.env.PATH = `${binDir}${path.delimiter}${originalPath}`;
    try {
      return await fn();
    } finally {
      process.env.PATH = originalPath;
    }
  });
}

test('probeYoutubeTracks prefers srv3 over vtt for automatic captions', async () => {
  await withFakeYtDlp(
    {
      id: 'abc123',
      title: 'Example',
      automatic_captions: {
        'ja-orig': [
          { ext: 'vtt', url: 'https://example.com/ja.vtt', name: 'Japanese auto' },
          { ext: 'srv3', url: 'https://example.com/ja.srv3', name: 'Japanese auto' },
        ],
      },
    },
    async () => {
      const result = await probeYoutubeTracks('https://www.youtube.com/watch?v=abc123');
      assert.equal(result.videoId, 'abc123');
      assert.equal(result.tracks[0]?.downloadUrl, 'https://example.com/ja.srv3');
      assert.equal(result.tracks[0]?.fileExtension, 'srv3');
    },
  );
});

test('probeYoutubeTracks keeps preferring srt for manual captions', async () => {
  await withFakeYtDlp(
    {
      id: 'abc123',
      title: 'Example',
      subtitles: {
        ja: [
          { ext: 'srv3', url: 'https://example.com/ja.srv3', name: 'Japanese manual' },
          { ext: 'srt', url: 'https://example.com/ja.srt', name: 'Japanese manual' },
        ],
      },
    },
    async () => {
      const result = await probeYoutubeTracks('https://www.youtube.com/watch?v=abc123');
      assert.equal(result.tracks[0]?.downloadUrl, 'https://example.com/ja.srt');
      assert.equal(result.tracks[0]?.fileExtension, 'srt');
    },
  );
});
