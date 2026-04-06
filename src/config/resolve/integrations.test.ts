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

test('resolveConfig parses configured mpv launch mode', () => {
  const { resolved, warnings } = resolveConfig({
    mpv: {
      launchMode: 'maximized',
    },
  });

  assert.equal(resolved.mpv.launchMode, 'maximized');
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

test('resolveConfig maps deprecated mpv fullscreen alias when launchMode is absent', () => {
  const { resolved, warnings } = resolveConfig({
    mpv: {
      startFullscreen: true,
    },
  });

  assert.equal(resolved.mpv.launchMode, 'fullscreen');
  assert.equal(warnings.length, 1);
  assert.deepEqual(warnings[0], {
    path: 'mpv.startFullscreen',
    value: true,
    fallback: 'fullscreen',
    message: 'Legacy key is deprecated; use mpv.launchMode',
  });
});

test('resolveConfig warns for invalid mpv launch mode and ignores deprecated alias fallback', () => {
  const { resolved, warnings } = resolveConfig({
    mpv: {
      launchMode: 'cinema' as never,
      startFullscreen: true,
    },
  });

  assert.equal(resolved.mpv.launchMode, 'normal');
  assert.equal(warnings.length, 1);
  assert.deepEqual(warnings[0], {
    path: 'mpv.launchMode',
    value: 'cinema',
    fallback: 'normal',
    message: "Expected one of: 'normal', 'maximized', 'fullscreen'.",
  });
});

test('resolveConfig warns for invalid deprecated mpv fullscreen alias type', () => {
  const { resolved, warnings } = resolveConfig({
    mpv: {
      startFullscreen: 'yes' as never,
    },
  });

  assert.equal(resolved.mpv.launchMode, 'normal');
  assert.equal(warnings.length, 1);
  assert.deepEqual(warnings[0], {
    path: 'mpv.startFullscreen',
    value: 'yes',
    fallback: 'normal',
    message: 'Expected boolean.',
  });
});

test('resolveConfig prefers launchMode over deprecated mpv fullscreen alias', () => {
  const { resolved, warnings } = resolveConfig({
    mpv: {
      launchMode: 'maximized',
      startFullscreen: true,
    },
  });

  assert.equal(resolved.mpv.launchMode, 'maximized');
  assert.deepEqual(warnings, []);
});
