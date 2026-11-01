import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { setTimeout as sleep } from 'node:timers/promises';

import { execPluginCommand, resolvePluginSpawnTarget } from './exec';

// The test runtime (bun or node) doubles as a portable child process.
const runtime = process.execPath;

test('pipes stdin to the child and collects stdout', async () => {
  const result = await execPluginCommand(
    runtime,
    ['-e', 'process.stdin.on("data", (d) => process.stdout.write(String(d).toUpperCase()))'],
    { stdin: '字幕 line' },
  );

  assert.equal(result.exitCode, 0);
  assert.equal(result.stdout, '字幕 LINE');
});

test('resolves with a non-zero exit code and stderr', async () => {
  const result = await execPluginCommand(runtime, [
    '-e',
    'process.stderr.write("bad things"); process.exit(3)',
  ]);

  assert.equal(result.exitCode, 3);
  assert.equal(result.stderr, 'bad things');
});

test('drops the AppImage library path from the child environment', async () => {
  const previous = process.env.LD_LIBRARY_PATH;
  process.env.LD_LIBRARY_PATH = '/tmp/.mount_SubMinXXXXXX/usr/lib';
  try {
    const result = await execPluginCommand(
      runtime,
      ['-e', 'process.stdout.write(String(process.env.LD_LIBRARY_PATH) + "|" + process.env.EXTRA)'],
      { env: { EXTRA: 'kept' } },
    );
    assert.equal(result.stdout, 'undefined|kept');
  } finally {
    if (previous === undefined) delete process.env.LD_LIBRARY_PATH;
    else process.env.LD_LIBRARY_PATH = previous;
  }
});

test('aborting kills the child and rejects with AbortError', async () => {
  const controller = new AbortController();
  const pending = execPluginCommand(runtime, ['-e', 'setTimeout(() => {}, 60_000)'], {
    signal: controller.signal,
  });
  setTimeout(() => controller.abort(), 50);

  await assert.rejects(pending, { name: 'AbortError' });
});

test('force-kills a child that ignores SIGTERM after the grace period', async (t) => {
  if (process.platform === 'win32') return t.skip('POSIX signals only');
  const pidFile = path.join(await mkdtemp(path.join(os.tmpdir(), 'plugin-exec-')), 'pid');
  const controller = new AbortController();
  const pending = execPluginCommand(
    runtime,
    [
      '-e',
      `process.on("SIGTERM", () => {}); require("node:fs").writeFileSync(${JSON.stringify(pidFile)}, String(process.pid)); setInterval(() => {}, 1000);`,
    ],
    { signal: controller.signal },
    100,
  );

  let pid = 0;
  for (let i = 0; i < 100 && !pid; i += 1) {
    pid = Number(await readFile(pidFile, 'utf8').catch(() => '0'));
    if (!pid) await sleep(20);
  }
  assert.ok(pid > 0, 'child never reported its pid');
  try {
    controller.abort();
    await assert.rejects(pending, { name: 'AbortError' });
    let alive = true;
    for (let i = 0; i < 100 && alive; i += 1) {
      await sleep(20);
      alive = isRunning(pid);
    }
    assert.equal(alive, false, 'child survived SIGTERM');
  } finally {
    if (isRunning(pid)) process.kill(pid, 'SIGKILL');
    await rm(path.dirname(pidFile), { recursive: true, force: true });
  }
});

function isRunning(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

test('rejects when the binary does not exist', async () => {
  await assert.rejects(
    execPluginCommand('subminer-definitely-missing-binary', []),
    (error: NodeJS.ErrnoException) => error.code === 'ENOENT',
  );
});

test('Windows batch shims run through cmd.exe; other commands spawn directly', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'plugin-exec-shim-'));
  try {
    await writeFile(path.join(dir, 'codex.cmd'), '');
    await writeFile(path.join(dir, 'native.exe'), '');
    const env = { PATH: dir, PATHEXT: '.exe;.cmd', ComSpec: 'cmd.exe' };

    const shim = resolvePluginSpawnTarget('codex', ['exec', '-'], env, 'win32');
    assert.equal(shim.viaShell, true);
    assert.equal(shim.file, 'cmd.exe');
    assert.deepEqual(shim.args.slice(0, 4), ['/d', '/s', '/v:off', '/c']);
    assert.ok(shim.args[4]!.includes(path.join(dir, 'codex.cmd')));

    assert.deepEqual(resolvePluginSpawnTarget('native', ['-v'], env, 'win32'), {
      file: path.join(dir, 'native.exe'),
      args: ['-v'],
      viaShell: false,
    });
    assert.deepEqual(resolvePluginSpawnTarget('codex', ['-'], env, 'linux'), {
      file: 'codex',
      args: ['-'],
      viaShell: false,
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
