import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

const {
  LINUX_FFMPEG_LIBRARY,
  stageLinuxAppImageSharedLibrary,
} = require('./electron-builder-after-pack.cjs') as {
  LINUX_FFMPEG_LIBRARY: string;
  stageLinuxAppImageSharedLibrary: (context: {
    appOutDir: string;
    electronPlatformName: string;
  }) => Promise<boolean>;
};

function createWorkspace(name: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), `${name}-`));
}

test('stageLinuxAppImageSharedLibrary copies libffmpeg.so into usr/lib for Linux packaging', async () => {
  const workspace = createWorkspace('subminer-after-pack-linux');
  const appOutDir = path.join(workspace, 'SubMiner-linux-x64');
  const sourceLibraryPath = path.join(appOutDir, LINUX_FFMPEG_LIBRARY);
  const targetLibraryPath = path.join(appOutDir, 'usr', 'lib', LINUX_FFMPEG_LIBRARY);

  fs.mkdirSync(appOutDir, { recursive: true });
  fs.writeFileSync(sourceLibraryPath, 'bundled ffmpeg', 'utf8');

  try {
    const staged = await stageLinuxAppImageSharedLibrary({
      appOutDir,
      electronPlatformName: 'linux',
    });

    assert.equal(staged, true);
    assert.equal(fs.readFileSync(targetLibraryPath, 'utf8'), 'bundled ffmpeg');
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('stageLinuxAppImageSharedLibrary skips non-Linux packaging contexts', async () => {
  const workspace = createWorkspace('subminer-after-pack-non-linux');
  const appOutDir = path.join(workspace, 'SubMiner-darwin-arm64');
  const sourceLibraryPath = path.join(appOutDir, LINUX_FFMPEG_LIBRARY);
  const targetLibraryPath = path.join(appOutDir, 'usr', 'lib', LINUX_FFMPEG_LIBRARY);

  fs.mkdirSync(appOutDir, { recursive: true });
  fs.writeFileSync(sourceLibraryPath, 'bundled ffmpeg', 'utf8');

  try {
    const staged = await stageLinuxAppImageSharedLibrary({
      appOutDir,
      electronPlatformName: 'darwin',
    });

    assert.equal(staged, false);
    assert.equal(fs.existsSync(targetLibraryPath), false);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('stageLinuxAppImageSharedLibrary no-ops when libffmpeg.so is absent', async () => {
  const workspace = createWorkspace('subminer-after-pack-missing-library');
  const appOutDir = path.join(workspace, 'SubMiner-linux-x64');
  const targetLibraryPath = path.join(appOutDir, 'usr', 'lib', LINUX_FFMPEG_LIBRARY);

  fs.mkdirSync(appOutDir, { recursive: true });

  try {
    const staged = await stageLinuxAppImageSharedLibrary({
      appOutDir,
      electronPlatformName: 'linux',
    });

    assert.equal(staged, false);
    assert.equal(fs.existsSync(targetLibraryPath), false);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});
