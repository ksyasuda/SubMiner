import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFile } from 'node:child_process';
import {
  APPIMAGE_MOUNT_KEEPALIVE_LABEL,
  APPIMAGE_MOUNT_KEEPALIVE_SCRIPT,
  resolveAppImageMountKeepaliveInvocation,
} from './appimage-mount-keepalive';

test('resolveAppImageMountKeepaliveInvocation is linux-only', () => {
  const env = { APPIMAGE: '/opt/SubMiner.AppImage' };
  assert.equal(resolveAppImageMountKeepaliveInvocation(env, 'win32'), null);
  assert.equal(resolveAppImageMountKeepaliveInvocation(env, 'darwin'), null);
  assert.notEqual(resolveAppImageMountKeepaliveInvocation(env, 'linux'), null);
});

test('resolveAppImageMountKeepaliveInvocation requires APPIMAGE env', () => {
  assert.equal(resolveAppImageMountKeepaliveInvocation({}, 'linux'), null);
  assert.equal(resolveAppImageMountKeepaliveInvocation({ APPIMAGE: '   ' }, 'linux'), null);
});

test('resolveAppImageMountKeepaliveInvocation honors disable env', () => {
  const env = {
    APPIMAGE: '/opt/SubMiner.AppImage',
    SUBMINER_NO_APPIMAGE_MOUNT_KEEPALIVE: '1',
  };
  assert.equal(resolveAppImageMountKeepaliveInvocation(env, 'linux'), null);
});

test('resolveAppImageMountKeepaliveInvocation builds sh invocation with AppImage path', () => {
  const invocation = resolveAppImageMountKeepaliveInvocation(
    { APPIMAGE: '/opt/SubMiner.AppImage' },
    'linux',
  );
  assert.ok(invocation);
  assert.equal(invocation.command, '/bin/sh');
  assert.deepEqual(invocation.args, [
    '-c',
    APPIMAGE_MOUNT_KEEPALIVE_SCRIPT,
    APPIMAGE_MOUNT_KEEPALIVE_LABEL,
    '/opt/SubMiner.AppImage',
  ]);
});

function runKeepaliveScript(
  appImagePath: string,
  extraArgs: string[] = [],
): Promise<{ status: number }> {
  return new Promise((resolve, reject) => {
    execFile(
      '/bin/sh',
      [
        '-c',
        APPIMAGE_MOUNT_KEEPALIVE_SCRIPT,
        APPIMAGE_MOUNT_KEEPALIVE_LABEL,
        appImagePath,
        ...extraArgs,
      ],
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
