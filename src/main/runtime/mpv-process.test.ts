import assert from 'node:assert/strict';
import { once } from 'node:events';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  createWindowsMpvPathDeps,
  getConfiguredWindowsMpvPathStatus,
  resolveMpvExecutablePath,
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
