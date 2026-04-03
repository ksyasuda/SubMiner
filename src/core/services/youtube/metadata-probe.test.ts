import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { probeYoutubeVideoMetadata } from './metadata-probe';

async function withTempDir<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-youtube-metadata-probe-'));
  try {
    return await fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function makeFakeYtDlpScript(dir: string, payload: string): void {
  const scriptPath = path.join(dir, 'yt-dlp');
  const script = process.platform === 'win32'
    ? `#!/usr/bin/env bun
process.stdout.write(${JSON.stringify(payload)});
`
    : `#!/usr/bin/env sh
cat <<'EOF' | base64 -d
${Buffer.from(payload).toString('base64')}
EOF
`;
  fs.writeFileSync(scriptPath, script, 'utf8');
  if (process.platform !== 'win32') {
    fs.chmodSync(scriptPath, 0o755);
  }
  fs.writeFileSync(scriptPath + '.cmd', `@echo off\r\nnode "${scriptPath}"\r\n`, 'utf8');
}

function makeHangingFakeYtDlpScript(dir: string): void {
  const scriptPath = path.join(dir, 'yt-dlp');
  const script = `#!/usr/bin/env sh
while :; do
  sleep 1;
done
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
    const originalPath = process.env.PATH ?? '';
    process.env.PATH = `${binDir}${path.delimiter}${originalPath}`;
    try {
      return await fn();
    } finally {
      process.env.PATH = originalPath;
    }
  });
}

async function withHangingFakeYtDlp<T>(fn: () => Promise<T>): Promise<T> {
  return await withTempDir(async (root) => {
    const binDir = path.join(root, 'bin');
    fs.mkdirSync(binDir, { recursive: true });
    makeHangingFakeYtDlpScript(binDir);
    const originalPath = process.env.PATH ?? '';
    process.env.PATH = `${binDir}${path.delimiter}${originalPath}`;
    try {
      return await fn();
    } finally {
      process.env.PATH = originalPath;
    }
  });
}

test('probeYoutubeVideoMetadata returns null on malformed yt-dlp JSON', async () => {
  await withFakeYtDlp('not-json', async () => {
    const result = await probeYoutubeVideoMetadata('https://www.youtube.com/watch?v=abc123');
    assert.equal(result, null);
  });
});

test('probeYoutubeVideoMetadata times out when yt-dlp hangs', { timeout: 20_000 }, async () => {
  await withHangingFakeYtDlp(async () => {
    await assert.rejects(
      probeYoutubeVideoMetadata('https://www.youtube.com/watch?v=abc123'),
      /timed out after 15000ms/,
    );
  });
});
