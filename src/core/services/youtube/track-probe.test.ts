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

function makeFakeYtDlpScript(dir: string, payload: unknown, rawScript = false): void {
  const scriptPath = path.join(dir, 'yt-dlp');
  const stdoutBody = typeof payload === 'string' ? payload : JSON.stringify(payload);
  const script =
    process.platform === 'win32'
      ? rawScript
        ? stdoutBody
        : `#!/usr/bin/env node
process.stdout.write(${JSON.stringify(stdoutBody)});
`
      : `#!/usr/bin/env sh
cat <<'EOF' | base64 -d
${Buffer.from(rawScript ? stdoutBody : `${JSON.stringify(stdoutBody)}`).toString('base64')}
EOF
`;
  fs.writeFileSync(scriptPath, script, 'utf8');
  if (process.platform !== 'win32') {
    fs.chmodSync(scriptPath, 0o755);
  }
  fs.writeFileSync(scriptPath + '.cmd', `@echo off\r\nnode "${scriptPath}"\r\n`, 'utf8');
}

async function withFakeYtDlp<T>(
  payload: unknown,
  fn: () => Promise<T>,
  options: { rawScript?: boolean } = {},
): Promise<T> {
  return await withTempDir(async (root) => {
    const binDir = path.join(root, 'bin');
    fs.mkdirSync(binDir, { recursive: true });
    makeFakeYtDlpScript(binDir, payload, options.rawScript === true);
    const originalPath = process.env.PATH ?? '';
    process.env.PATH = `${binDir}${path.delimiter}${originalPath}`;
    try {
      return await fn();
    } finally {
      process.env.PATH = originalPath;
    }
  });
}

async function withFakeYtDlpCommand<T>(
  payload: unknown,
  fn: () => Promise<T>,
  options: { rawScript?: boolean } = {},
): Promise<T> {
  return await withTempDir(async (root) => {
    const binDir = path.join(root, 'bin');
    fs.mkdirSync(binDir, { recursive: true });
    makeFakeYtDlpScript(binDir, payload, options.rawScript === true);
    const originalPath = process.env.PATH;
    const originalCommand = process.env.SUBMINER_YTDLP_BIN;
    process.env.PATH = '';
    process.env.SUBMINER_YTDLP_BIN =
      process.platform === 'win32' ? path.join(binDir, 'yt-dlp.cmd') : path.join(binDir, 'yt-dlp');
    try {
      return await fn();
    } finally {
      if (originalPath === undefined) {
        delete process.env.PATH;
      } else {
        process.env.PATH = originalPath;
      }
      if (originalCommand === undefined) {
        delete process.env.SUBMINER_YTDLP_BIN;
      } else {
        process.env.SUBMINER_YTDLP_BIN = originalCommand;
      }
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

test('probeYoutubeTracks honors SUBMINER_YTDLP_BIN when yt-dlp is not on PATH', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await withFakeYtDlpCommand(
    {
      id: 'abc123',
      title: 'Example',
      subtitles: {
        ja: [{ ext: 'vtt', url: 'https://example.com/ja.vtt', name: 'Japanese manual' }],
      },
    },
    async () => {
      const result = await probeYoutubeTracks('https://www.youtube.com/watch?v=abc123');
      assert.equal(result.videoId, 'abc123');
      assert.equal(result.tracks[0]?.downloadUrl, 'https://example.com/ja.vtt');
      assert.equal(result.tracks[0]?.fileExtension, 'vtt');
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

test('probeYoutubeTracks reports malformed yt-dlp JSON with context', async () => {
  await withFakeYtDlp('not-json', async () => {
    await assert.rejects(
      async () => await probeYoutubeTracks('https://www.youtube.com/watch?v=abc123'),
      /Failed to parse yt-dlp output as JSON/,
    );
  });
});
