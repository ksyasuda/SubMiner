import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

function withTempDir<T>(fn: (dir: string) => T): T {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-mkv-to-readme-video-test-'));
  try {
    return fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function writeExecutable(filePath: string, contents: string): void {
  fs.writeFileSync(filePath, contents, 'utf8');
  fs.chmodSync(filePath, 0o755);
}

test('mkv-to-readme-video accepts libwebp_anim when libwebp is unavailable', () => {
  withTempDir((root) => {
    const binDir = path.join(root, 'bin');
    const inputPath = path.join(root, 'sample.mkv');
    const ffmpegLogPath = path.join(root, 'ffmpeg-args.log');

    fs.mkdirSync(binDir, { recursive: true });
    fs.writeFileSync(inputPath, 'fake-video', 'utf8');

    writeExecutable(
      path.join(binDir, 'ffmpeg'),
      `#!/usr/bin/env bash
set -euo pipefail

if [[ "$#" -ge 2 && "$1" == "-hide_banner" && "$2" == "-encoders" ]]; then
  cat <<'EOF'
 V....D libx264              libx264 H.264 / AVC / MPEG-4 AVC / MPEG-4 part 10 (codec h264)
 V....D libvpx-vp9           libvpx VP9 (codec vp9)
 V....D libwebp_anim         libwebp WebP image (codec webp)
 A....D aac                  AAC (Advanced Audio Coding)
 A....D libopus              libopus Opus (codec opus)
EOF
  exit 0
fi

printf '%s\\n' "$*" >> "${ffmpegLogPath}"
output=""
for arg in "$@"; do
  output="$arg"
done
mkdir -p "$(dirname "$output")"
touch "$output"
`,
    );

    const result = spawnSync('bash', ['scripts/mkv-to-readme-video.sh', '--webp', inputPath], {
      cwd: process.cwd(),
      env: {
        ...process.env,
        PATH: `${binDir}:${process.env.PATH || ''}`,
      },
      encoding: 'utf8',
    });

    assert.equal(result.status, 0, result.stderr || result.stdout);
    assert.equal(fs.existsSync(path.join(root, 'sample.mp4')), true);
    assert.equal(fs.existsSync(path.join(root, 'sample.webm')), true);
    assert.equal(fs.existsSync(path.join(root, 'sample.webp')), true);
    assert.equal(fs.existsSync(path.join(root, 'sample-poster.jpg')), true);

    const ffmpegLog = fs.readFileSync(ffmpegLogPath, 'utf8');
    assert.match(ffmpegLog, /-c:v libwebp_anim/);
  });
});
