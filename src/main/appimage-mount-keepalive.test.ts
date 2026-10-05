import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFile } from 'node:child_process';
import {
  resolveAppImageMountKeepaliveInvocation,
  resolveMountFsType,
} from './appimage-mount-keepalive';

const FUSE_EXEC_PATH = '/tmp/.mount_SubMinAb12Cd/SubMiner';
const MOUNT_INFO = [
  '1 0 0:1 / / rw,relatime - btrfs /dev/nvme0n1p2 rw',
  '773 75 0:95 / /tmp/.mount_SubMinAb12Cd ro,nosuid shared:809 - fuse.SubMiner.AppImage SubMiner.AppImage ro',
  '1238 1154 7:5 / /run/firejail/appimage ro,nosuid - squashfs /dev/loop5 ro',
  '900 75 0:99 / /mnt/with\\040space rw - ext4 /dev/sdb1 rw',
].join('\n');

function resolveOnFuse(env: NodeJS.ProcessEnv, platform: NodeJS.Platform) {
  return resolveAppImageMountKeepaliveInvocation(env, platform, FUSE_EXEC_PATH, () => MOUNT_INFO);
}

test('resolveAppImageMountKeepaliveInvocation is linux-only', () => {
  const env = { APPIMAGE: '/opt/SubMiner.AppImage' };
  assert.equal(resolveOnFuse(env, 'win32'), null);
  assert.equal(resolveOnFuse(env, 'darwin'), null);
  assert.notEqual(resolveOnFuse(env, 'linux'), null);
});

test('resolveAppImageMountKeepaliveInvocation requires APPIMAGE env', () => {
  assert.equal(resolveOnFuse({}, 'linux'), null);
  assert.equal(resolveOnFuse({ APPIMAGE: '   ' }, 'linux'), null);
});

test('resolveAppImageMountKeepaliveInvocation honors disable env', () => {
  const env = {
    APPIMAGE: '/opt/SubMiner.AppImage',
    SUBMINER_NO_APPIMAGE_MOUNT_KEEPALIVE: '1',
  };
  assert.equal(resolveOnFuse(env, 'linux'), null);
});

test('resolveAppImageMountKeepaliveInvocation skips kernel squashfs mounts owned by a sandbox', () => {
  const env = { APPIMAGE: '/tmp/SubMiner.AppImage' };
  assert.equal(
    resolveAppImageMountKeepaliveInvocation(
      env,
      'linux',
      '/run/firejail/appimage/SubMiner',
      () => MOUNT_INFO,
    ),
    null,
  );
  assert.notEqual(
    resolveAppImageMountKeepaliveInvocation(
      env,
      'linux',
      '/tmp/appimage_extracted_42c4346b/SubMiner',
      () => MOUNT_INFO,
    ),
    null,
    'extract-and-run directories are deleted with the bootstrap, so they keep the supervisor',
  );
  assert.notEqual(
    resolveAppImageMountKeepaliveInvocation(env, 'linux', FUSE_EXEC_PATH, () => null),
    null,
    'unknown mount type keeps the supervisor',
  );
});

test('resolveMountFsType picks the longest containing mount point', () => {
  assert.equal(resolveMountFsType(FUSE_EXEC_PATH, MOUNT_INFO), 'fuse.SubMiner.AppImage');
  assert.equal(resolveMountFsType('/run/firejail/appimage/SubMiner', MOUNT_INFO), 'squashfs');
  assert.equal(resolveMountFsType('/mnt/with space/SubMiner', MOUNT_INFO), 'ext4');
  assert.equal(resolveMountFsType('/tmp/.mount_SubMinAb12CdX/SubMiner', MOUNT_INFO), 'btrfs');
  assert.equal(resolveMountFsType('/usr/bin/sh', ''), null);
});

function runKeepaliveScript(
  appImagePath: string,
  extraArgs: string[] = [],
): Promise<{ status: number }> {
  // Run the exact invocation the app would spawn, so the argv is exercised too.
  const invocation = resolveOnFuse({ APPIMAGE: appImagePath }, 'linux');
  assert.ok(invocation);
  return new Promise((resolve, reject) => {
    execFile(
      invocation.command,
      [...invocation.args, ...extraArgs],
      { timeout: 30_000 },
      (error) => {
        if (error && typeof error.code !== 'number') {
          reject(error);
          return;
        }
        resolve({ status: typeof error?.code === 'number' ? error.code : 0 });
      },
    );
  });
}

function writeExecutable(filePath: string, content: string): void {
  fs.writeFileSync(filePath, content, { mode: 0o755 });
}

const linuxTest = process.platform === 'linux' ? test : test.skip;

linuxTest('keepalive script releases the mount only after straggler processes exit', async () => {
  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-keepalive-test-'));
  const resultsDir = path.join(workDir, 'results');
  fs.mkdirSync(resultsDir);
  const mountDir = path.join(workDir, 'fake-mount');
  fs.mkdirSync(mountDir);

  // AppRun leaves behind a straggler that keeps executing *from the mount*
  // after AppRun itself exits — mimicking Chromium utility children.
  fs.copyFileSync('/usr/bin/sleep', path.join(mountDir, 'straggler'));
  fs.chmodSync(path.join(mountDir, 'straggler'), 0o755);
  writeExecutable(
    path.join(mountDir, 'AppRun'),
    [
      '#!/bin/sh',
      `"${mountDir}/straggler" 1 &`,
      `date +%s%N > "${resultsDir}/apprun-exited"`,
      'exit 42',
    ].join('\n'),
  );

  const fakeAppImage = path.join(workDir, 'Fake.AppImage');
  writeExecutable(
    fakeAppImage,
    [
      '#!/bin/sh',
      'if [ "${1:-}" = "--appimage-mount" ]; then',
      `  echo "${mountDir}"`,
      `  trap ': > "${resultsDir}/holder-released"; sleep 0.1; date +%s%N > "${resultsDir}/holder-released"; exit 0' TERM INT`,
      '  while :; do sleep 0.05; done',
      'fi',
      `date +%s%N > "${resultsDir}/direct-run"`,
      'exit 0',
    ].join('\n'),
  );

  try {
    const { status } = await runKeepaliveScript(fakeAppImage);

    assert.equal(status, 42, 'exit code of AppRun must be propagated');
    assert.ok(
      !fs.existsSync(path.join(resultsDir, 'direct-run')),
      'must not fall back to direct AppImage run when mount succeeds',
    );
    // The script does not wait for the holder to finish handling SIGTERM
    // (the real runtime unmounts on its own after the signal), so poll.
    const releasedMarker = path.join(resultsDir, 'holder-released');
    const pollDeadline = Date.now() + 2000;
    let holderReleased: number | null = null;
    while (holderReleased === null && Date.now() < pollDeadline) {
      if (fs.existsSync(releasedMarker)) {
        const timestamp = fs.readFileSync(releasedMarker, 'utf8').trim();
        if (/^\d+$/.test(timestamp)) holderReleased = Number(timestamp);
      }
      if (holderReleased !== null) break;
      await new Promise((r) => setTimeout(r, 25));
    }
    assert.ok(holderReleased !== null, 'holder release timestamp must be recorded');

    const appRunExited = Number(
      fs.readFileSync(path.join(resultsDir, 'apprun-exited'), 'utf8').trim(),
    );
    const drainNs = holderReleased - appRunExited;
    assert.ok(
      drainNs >= 0.8e9,
      `holder must outlive the 1s straggler (drained after ${drainNs / 1e9}s)`,
    );
  } finally {
    fs.rmSync(workDir, { recursive: true, force: true });
  }
});

linuxTest('keepalive script falls back to direct run when --appimage-mount fails', async () => {
  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-keepalive-test-'));
  const resultsDir = path.join(workDir, 'results');
  fs.mkdirSync(resultsDir);

  const fakeAppImage = path.join(workDir, 'Fake.AppImage');
  writeExecutable(
    fakeAppImage,
    [
      '#!/bin/sh',
      'if [ "${1:-}" = "--appimage-mount" ]; then',
      '  exit 1',
      'fi',
      `printf '%s\\n' "$@" > "${resultsDir}/direct-run"`,
      'exit 7',
    ].join('\n'),
  );

  try {
    const { status } = await runKeepaliveScript(fakeAppImage, ['--start', '--background']);
    assert.equal(status, 7, 'direct-run exit code must be propagated');
    assert.equal(
      fs.readFileSync(path.join(resultsDir, 'direct-run'), 'utf8'),
      '--start\n--background\n',
      'launch args must be forwarded to the direct run',
    );
  } finally {
    fs.rmSync(workDir, { recursive: true, force: true });
  }
});
