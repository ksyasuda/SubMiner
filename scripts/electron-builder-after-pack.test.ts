import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

const {
  LINUX_FFMPEG_LIBRARY,
  MACOS_HELPERS,
  default: afterPack,
  stageLinuxAppImageSharedLibrary,
  verifyMacOSHelpers,
} = require('./electron-builder-after-pack.cjs') as {
  LINUX_FFMPEG_LIBRARY: string;
  MACOS_HELPERS: string[];
  default: (
    context: {
      appOutDir: string;
      arch?: number;
      electronPlatformName: string;
      packager?: { appInfo?: { productFilename?: string } };
    },
    deps?: {
      auditPackage?: (context: { appOutDir: string }) => Promise<void>;
      stageBunRuntime?: (options: {
        appOutDir: string;
        platform: string;
        arch: number | undefined;
        productFilename: string;
      }) => Promise<void>;
    },
  ) => Promise<void>;
  stageLinuxAppImageSharedLibrary: (context: {
    appOutDir: string;
    electronPlatformName: string;
  }) => Promise<boolean>;
  verifyMacOSHelpers: (context: {
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

test('verifyMacOSHelpers accepts packaged macOS helper binaries', async () => {
  const workspace = createWorkspace('subminer-after-pack-macos-helper');
  const appOutDir = path.join(workspace, 'SubMiner-darwin-arm64');
  const scriptsDir = path.join(appOutDir, 'SubMiner.app', 'Contents', 'Resources', 'scripts');

  fs.mkdirSync(scriptsDir, { recursive: true });
  for (const helperName of MACOS_HELPERS) {
    fs.writeFileSync(path.join(scriptsDir, helperName), 'compiled helper', 'utf8');
    fs.chmodSync(path.join(scriptsDir, helperName), 0o755);
  }

  try {
    const verified = await verifyMacOSHelpers({
      appOutDir,
      electronPlatformName: 'darwin',
    });

    assert.equal(verified, true);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('verifyMacOSHelpers throws when a macOS helper binary is missing', async () => {
  const workspace = createWorkspace('subminer-after-pack-macos-helper-missing');
  const appOutDir = path.join(workspace, 'SubMiner-darwin-arm64');
  const scriptsDir = path.join(appOutDir, 'SubMiner.app', 'Contents', 'Resources', 'scripts');

  fs.mkdirSync(scriptsDir, { recursive: true });
  fs.writeFileSync(path.join(scriptsDir, 'get-mpv-window-macos'), 'compiled helper', 'utf8');
  fs.chmodSync(path.join(scriptsDir, 'get-mpv-window-macos'), 0o755);

  try {
    await assert.rejects(
      verifyMacOSHelpers({
        appOutDir,
        electronPlatformName: 'darwin',
      }),
      /macOS packaging requires macos-controller-watch/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('stageLinuxAppImageSharedLibrary throws when Linux packaging is missing libffmpeg.so', async () => {
  const workspace = createWorkspace('subminer-after-pack-missing-library');
  const appOutDir = path.join(workspace, 'SubMiner-linux-x64');

  fs.mkdirSync(appOutDir, { recursive: true });

  try {
    await assert.rejects(
      stageLinuxAppImageSharedLibrary({
        appOutDir,
        electronPlatformName: 'linux',
      }),
      new RegExp(`Linux packaging requires ${LINUX_FFMPEG_LIBRARY} at .*${LINUX_FFMPEG_LIBRARY}`),
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('afterPack propagates Linux staging failures', async () => {
  const workspace = createWorkspace('subminer-after-pack-propagates-linux-failure');
  const appOutDir = path.join(workspace, 'SubMiner-linux-x64');

  fs.mkdirSync(appOutDir, { recursive: true });

  try {
    await assert.rejects(
      afterPack({
        appOutDir,
        electronPlatformName: 'linux',
      }),
      /Linux packaging requires libffmpeg\.so/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('afterPack stages Linux and Bun runtime assets before auditing the package', async () => {
  const workspace = createWorkspace('subminer-after-pack-target');
  const appOutDir = path.join(workspace, 'SubMiner-linux-arm64');
  const sourceLibraryPath = path.join(appOutDir, LINUX_FFMPEG_LIBRARY);
  const targetLibraryPath = path.join(appOutDir, 'usr', 'lib', LINUX_FFMPEG_LIBRARY);
  const operations: string[] = [];
  let stagedOptions:
    | {
        appOutDir: string;
        platform: string;
        arch: number | undefined;
        productFilename: string;
      }
    | undefined;

  fs.mkdirSync(appOutDir, { recursive: true });
  fs.writeFileSync(sourceLibraryPath, 'bundled ffmpeg', 'utf8');

  try {
    await afterPack(
      {
        appOutDir,
        arch: 3,
        electronPlatformName: 'linux',
        packager: { appInfo: { productFilename: 'SubMiner Preview' } },
      },
      {
        stageBunRuntime: async (options) => {
          stagedOptions = options;
          operations.push('stage-bun');
        },
        auditPackage: async (context) => {
          assert.equal(context.appOutDir, appOutDir);
          assert.equal(fs.readFileSync(targetLibraryPath, 'utf8'), 'bundled ffmpeg');
          operations.push('audit');
        },
      },
    );

    assert.deepEqual(operations, ['stage-bun', 'audit']);
    assert.deepEqual(stagedOptions, {
      appOutDir,
      platform: 'linux',
      arch: 3,
      productFilename: 'SubMiner Preview',
    });
    assert.equal(fs.readFileSync(targetLibraryPath, 'utf8'), 'bundled ffmpeg');
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});
