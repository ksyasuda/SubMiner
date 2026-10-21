import assert from 'node:assert/strict';
import { once } from 'node:events';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  createPosixMpvPathDeps,
  createWindowsMpvPathDeps,
  getConfiguredWindowsMpvPathStatus,
  resolveMpvExecutablePath,
  resolvePosixMpvPath,
  resolveWindowsMpvPath,
  spawnMpvProcess,
} from './mpv-process';

const testWindows = process.platform === 'win32' ? test : test.skip;
const testLinux = process.platform === 'linux' ? test : test.skip;

function makeWindowsDeps(options: {
  existing?: string[];
  env?: Record<string, string>;
  where?: { status: number | null; stdout: string };
}) {
  const existing = new Set(options.existing ?? []);
  const calls: string[] = [];
  const deps = createWindowsMpvPathDeps({
    getEnv: (name) => options.env?.[name],
    fileExists: (candidate) => existing.has(candidate),
    runWhere: () => {
      calls.push('where');
      return options.where ?? { status: 1, stdout: '' };
    },
  });
  return { deps, calls };
}

test('getConfiguredWindowsMpvPathStatus classifies blank, configured and invalid paths', () => {
  const exists = (candidate: string) => candidate === 'C:\\mpv\\mpv.exe';

  assert.equal(getConfiguredWindowsMpvPathStatus('   ', exists), 'blank');
  assert.equal(getConfiguredWindowsMpvPathStatus('  C:\\mpv\\mpv.exe ', exists), 'configured');
  assert.equal(getConfiguredWindowsMpvPathStatus('C:\\missing\\mpv.exe', exists), 'invalid');
});

test('resolveWindowsMpvPath prefers a valid configured path and trims it', () => {
  const { deps, calls } = makeWindowsDeps({
    existing: ['C:\\cfg\\mpv.exe', 'C:\\env\\mpv.exe'],
    env: { SUBMINER_MPV_PATH: 'C:\\env\\mpv.exe' },
  });

  assert.equal(resolveWindowsMpvPath(deps, '  C:\\cfg\\mpv.exe  '), 'C:\\cfg\\mpv.exe');
  assert.deepEqual(calls, []);
});

test('resolveWindowsMpvPath does not fall back when the configured path is invalid', () => {
  const { deps, calls } = makeWindowsDeps({
    existing: ['C:\\env\\mpv.exe'],
    env: { SUBMINER_MPV_PATH: 'C:\\env\\mpv.exe' },
    where: { status: 0, stdout: 'C:\\env\\mpv.exe\r\n' },
  });

  assert.equal(resolveWindowsMpvPath(deps, 'C:\\missing\\mpv.exe'), '');
  assert.deepEqual(calls, []);
});

test('resolveWindowsMpvPath uses SUBMINER_MPV_PATH before searching PATH', () => {
  const { deps, calls } = makeWindowsDeps({
    existing: ['C:\\env\\mpv.exe'],
    env: { SUBMINER_MPV_PATH: ' C:\\env\\mpv.exe ' },
  });

  assert.equal(resolveWindowsMpvPath(deps), 'C:\\env\\mpv.exe');
  assert.deepEqual(calls, []);
});

test('resolveWindowsMpvPath takes the first where.exe line that exists', () => {
  const { deps } = makeWindowsDeps({
    existing: ['C:\\second\\mpv.exe', 'C:\\third\\mpv.exe'],
    env: { SUBMINER_MPV_PATH: 'C:\\stale\\mpv.exe' },
    where: {
      status: 0,
      stdout: '\r\nC:\\first\\mpv.exe\r\n  C:\\second\\mpv.exe  \r\nC:\\third\\mpv.exe\r\n',
    },
  });

  assert.equal(resolveWindowsMpvPath(deps), 'C:\\second\\mpv.exe');
});

test('resolveWindowsMpvPath returns blank when nothing resolves', () => {
  const failed = makeWindowsDeps({ where: { status: 1, stdout: 'C:\\first\\mpv.exe' } });
  const blank = makeWindowsDeps({ where: { status: 0, stdout: '\r\n  \r\n' } });
  const missing = makeWindowsDeps({ where: { status: 0, stdout: 'C:\\first\\mpv.exe' } });

  assert.equal(resolveWindowsMpvPath(failed.deps), '');
  assert.equal(resolveWindowsMpvPath(blank.deps), '');
  assert.equal(resolveWindowsMpvPath(missing.deps), '');
});

// launchd's default PATH, which apps opened from Finder, the Dock, or Spotlight inherit.
const LAUNCHD_PATH = '/usr/bin:/bin:/usr/sbin:/sbin';

function makePosixDeps(options: {
  platform: NodeJS.Platform;
  executables?: string[];
  env?: Record<string, string>;
}) {
  const executables = new Set(options.executables ?? []);
  return createPosixMpvPathDeps({
    platform: options.platform,
    getEnv: (name) => options.env?.[name],
    isExecutable: (candidate) => executables.has(candidate),
  });
}

test('resolvePosixMpvPath prefers a configured executable and does not fall back when it is invalid', () => {
  const deps = makePosixDeps({
    platform: 'darwin',
    executables: ['/cfg/mpv', '/opt/homebrew/bin/mpv'],
    env: { PATH: LAUNCHD_PATH },
  });

  assert.equal(resolvePosixMpvPath(deps, '  /cfg/mpv  '), '/cfg/mpv');
  assert.equal(resolvePosixMpvPath(deps, '/missing/mpv'), '');
});

test('resolvePosixMpvPath uses SUBMINER_MPV_PATH before searching PATH', () => {
  const deps = makePosixDeps({
    platform: 'linux',
    executables: ['/env/mpv', '/usr/bin/mpv'],
    env: { SUBMINER_MPV_PATH: ' /env/mpv ', PATH: '/usr/bin' },
  });

  assert.equal(resolvePosixMpvPath(deps), '/env/mpv');
});

test('resolvePosixMpvPath takes the first PATH entry that has mpv', () => {
  const deps = makePosixDeps({
    platform: 'darwin',
    executables: ['/second/mpv', '/third/mpv', '/opt/homebrew/bin/mpv'],
    env: { SUBMINER_MPV_PATH: '/stale/mpv', PATH: '/first::/second:/third' },
  });

  assert.equal(resolvePosixMpvPath(deps), '/second/mpv');
});

test('resolvePosixMpvPath finds Homebrew mpv on macOS when the app only has the launchd PATH', () => {
  const homebrew = makePosixDeps({
    platform: 'darwin',
    executables: ['/opt/homebrew/bin/mpv'],
    env: { PATH: LAUNCHD_PATH },
  });
  const appBundle = makePosixDeps({
    platform: 'darwin',
    executables: ['/Applications/mpv.app/Contents/MacOS/mpv'],
    env: { PATH: LAUNCHD_PATH },
  });

  assert.equal(resolvePosixMpvPath(homebrew), '/opt/homebrew/bin/mpv');
  assert.equal(resolvePosixMpvPath(appBundle), '/Applications/mpv.app/Contents/MacOS/mpv');
});

test('resolvePosixMpvPath only probes macOS install locations on macOS', () => {
  const deps = makePosixDeps({
    platform: 'linux',
    executables: ['/opt/homebrew/bin/mpv'],
    env: { PATH: LAUNCHD_PATH },
  });

  assert.equal(resolvePosixMpvPath(deps), '');
});

const X11_SPAWN_CASES = [
  {
    name: 'forces the X11 backend for an unsupported Wayland session',
    compositorEnv: {},
    forcesX11: true,
  },
  {
    name: 'leaves a natively supported compositor session alone',
    compositorEnv: { HYPRLAND_INSTANCE_SIGNATURE: 'sig' },
    forcesX11: false,
  },
];

for (const c of X11_SPAWN_CASES) {
  testLinux(`spawnMpvProcess ${c.name}`, async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-mpv-process-'));
    // The child reports what it received so args and env can be asserted from outside.
    const scriptPath = path.join(dir, 'probe.js');
    const reportPath = path.join(dir, 'report.json');
    fs.writeFileSync(
      scriptPath,
      `require('node:fs').writeFileSync(process.argv[2], JSON.stringify({
        args: process.argv.slice(3),
        wayland: process.env.WAYLAND_DISPLAY ?? null,
        sessionType: process.env.XDG_SESSION_TYPE ?? null,
      }));`,
    );
    const child = spawnMpvProcess(process.execPath, [scriptPath, reportPath, '--user-arg'], {
      DISPLAY: ':0',
      WAYLAND_DISPLAY: 'wayland-1',
      XDG_SESSION_TYPE: 'wayland',
      ...c.compositorEnv,
    });
    try {
      await once(child, 'exit');
      const report = JSON.parse(fs.readFileSync(reportPath, 'utf8'));
      if (c.forcesX11) {
        assert.deepEqual(report, {
          args: ['--user-arg', '--gpu-context=x11vk,x11egl,x11'],
          wayland: null,
          sessionType: 'x11',
        });
      } else {
        assert.deepEqual(report, {
          args: ['--user-arg'],
          wayland: 'wayland-1',
          sessionType: 'wayland',
        });
      }
    } finally {
      if (child.exitCode === null) child.kill();
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
}

testWindows('managed Windows playback launches the configured executable', async () => {
  const executable = resolveMpvExecutablePath(`  ${process.execPath}  `);
  const child = spawnMpvProcess(executable, ['-e', 'process.exit(17)']);
  try {
    assert.equal(child.spawnfile, process.execPath);
    const [code] = await once(child, 'exit');
    assert.equal(code, 17);
  } finally {
    if (child.exitCode === null) child.kill();
  }
});

testWindows('managed Windows playback rejects an invalid configured executable', () => {
  assert.throws(
    () => resolveMpvExecutablePath(`${process.execPath}/missing-mpv.exe`),
    /Could not find mpv.exe/,
  );
});
