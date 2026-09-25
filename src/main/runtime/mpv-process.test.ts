import assert from 'node:assert/strict';
import { once } from 'node:events';
import test from 'node:test';
import { resolveMpvExecutablePath, spawnMpvProcess } from './mpv-process';

const testWindows = process.platform === 'win32' ? test : test.skip;

test('mpv process launcher forwards arguments and environment to the child', async () => {
  const child = spawnMpvProcess(
    process.execPath,
    ['-e', 'process.exit(Number(process.env.SUBMINER_TEST_EXIT))'],
    { ...process.env, DISPLAY: '', SUBMINER_TEST_EXIT: '17' },
  );
  try {
    const [code] = await once(child, 'exit');
    assert.equal(code, 17);
  } finally {
    if (child.exitCode === null) child.kill();
  }
});

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
