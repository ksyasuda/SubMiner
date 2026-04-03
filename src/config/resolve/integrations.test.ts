import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveConfig } from '../resolve';

test('resolveConfig trims configured mpv executable path', () => {
  const { resolved, warnings } = resolveConfig({
    mpv: {
      executablePath: '  C:\\Program Files\\mpv\\mpv.exe  ',
    },
  });

  assert.equal(resolved.mpv.executablePath, 'C:\\Program Files\\mpv\\mpv.exe');
  assert.deepEqual(warnings, []);
});

test('resolveConfig warns for invalid mpv executable path type', () => {
  const { resolved, warnings } = resolveConfig({
    mpv: {
      executablePath: 42 as never,
    },
  });

  assert.equal(resolved.mpv.executablePath, '');
  assert.equal(warnings.length, 1);
  assert.deepEqual(warnings[0], {
    path: 'mpv.executablePath',
    value: 42,
    fallback: '',
    message: 'Expected string.',
  });
});
