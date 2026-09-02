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

function shellQuote(value: string): string {
  return `'${value.replace(/'/g, `'\"'\"'`)}'`;
}

function toBashPath(filePath: string): string {
  if (process.platform !== 'win32') return filePath;

  const normalized = filePath.replace(/\\/g, '/');
  const match = normalized.match(/^([A-Za-z]):\/(.*)$/);
  if (!match) return normalized;

  const drive = match[1]!;
  const rest = match[2]!;
  const probe = spawnSync('bash', ['-c', 'uname -s'], { encoding: 'utf8' });
  if (probe.status === 0 && /linux/i.test(probe.stdout)) {
    return `/mnt/${drive.toLowerCase()}/${rest}`;
  }

  return `${drive.toUpperCase()}:/${rest}`;
}

test('mkv-to-readme-video builds every output with the scaled mpv crop', () => {
  withTempDir((root) => {
    const binDir = path.join(root, 'bin');
    const inputPath = path.join(root, 'sample.mkv');
    const ffmpegLogPath = path.join(root, 'ffmpeg-args.log');
    const ffmpegLogPathBash = toBashPath(ffmpegLogPath);

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

if [[ "$#" -eq 0 ]]; then
  exit 0
fi

printf '%s\\n' "$*" >> "${ffmpegLogPathBash}"
output=""
for arg in "$@"; do
  output="$arg"
done
if [[ -z "$output" ]]; then
  exit 0
fi
mkdir -p "$(dirname "$output")"
touch "$output"
`,
    );

    const ffmpegShimPath = toBashPath(path.join(binDir, 'ffmpeg'));
    const ffmpegShimDir = toBashPath(binDir);
    const inputBashPath = toBashPath(inputPath);
    const command = [
      `chmod +x ${shellQuote(ffmpegShimPath)}`,
      `PATH=${shellQuote(`${ffmpegShimDir}:`)}"$PATH"`,
      `scripts/mkv-to-readme-video.sh --webp ${shellQuote(inputBashPath)}`,
    ].join('; ');
    const result = spawnSync('bash', ['-lc', command], {
      cwd: process.cwd(),
      encoding: 'utf8',
    });

    assert.equal(result.status, 0, result.stderr || result.stdout);
    assert.equal(fs.existsSync(path.join(root, 'sample.mp4')), true);
    assert.equal(fs.existsSync(path.join(root, 'sample.webm')), true);
    assert.equal(fs.existsSync(path.join(root, 'sample.webp')), true);
    assert.equal(fs.existsSync(path.join(root, 'sample-poster.jpg')), true);

    const ffmpegLog = fs.readFileSync(ffmpegLogPath, 'utf8');
    assert.match(ffmpegLog, /-c:v libwebp_anim/);
    const scaledCropUses = ffmpegLog.match(
      /-vf crop=1920\*iw\/3440:1080\*ih\/1440:760\*iw\/3440:205\*ih\/1440,scale=1920:1080:flags=lanczos/g,
    );
    assert.equal(scaledCropUses?.length, 4);
  });
});
