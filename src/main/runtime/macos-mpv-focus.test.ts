import assert from 'node:assert/strict';
import test from 'node:test';
import { focusMacOSMpvProcess } from './macos-mpv-focus';

test('focusMacOSMpvProcess fronts the running mpv process with osascript', async () => {
  const calls: Array<{ command: string; args: string[]; timeout?: number }> = [];

  await focusMacOSMpvProcess({
    execFile: (command, args, options, callback) => {
      calls.push({ command, args, timeout: options.timeout });
      callback(null);
    },
  });

  assert.equal(calls.length, 1);
  assert.equal(calls[0]?.command, '/usr/bin/osascript');
  assert.equal(calls[0]?.timeout, 2000);
  assert.deepEqual(calls[0]?.args, [
    '-e',
    'tell application "System Events" to set frontmost of the first process whose name is "mpv" to true',
  ]);
});

test('focusMacOSMpvProcess resolves when osascript fails', async () => {
  await focusMacOSMpvProcess({
    execFile: (_command, _args, _options, callback) => {
      callback(new Error('not allowed'));
    },
  });
});
