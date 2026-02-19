import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { resolveConfigBaseDirs, resolveConfigDir, resolveConfigFilePath } from './path-resolution';

function existsSyncFrom(paths: string[]): (candidate: string) => boolean {
  const normalized = new Set(paths.map((entry) => path.normalize(entry)));
  return (candidate: string): boolean => normalized.has(path.normalize(candidate));
}

test('resolveConfigBaseDirs trims xdg value and deduplicates fallback dir', () => {
  const homeDir = '/home/tester';
  const baseDirs = resolveConfigBaseDirs(' /home/tester/.config ', homeDir);
  assert.deepEqual(baseDirs, [path.join(homeDir, '.config')]);
});

test('resolveConfigDir prefers xdg SubMiner config when present', () => {
  const homeDir = '/home/tester';
  const xdgConfigHome = '/tmp/xdg-config';
  const configDir = path.join(xdgConfigHome, 'SubMiner');
  const existsSync = existsSyncFrom([path.join(configDir, 'config.jsonc')]);

  const resolved = resolveConfigDir({
    xdgConfigHome,
    homeDir,
    existsSync,
  });

  assert.equal(resolved, configDir);
});

test('resolveConfigDir falls back to lowercase subminer candidate', () => {
  const homeDir = '/home/tester';
  const configDir = path.join(homeDir, '.config', 'subminer');
  const existsSync = existsSyncFrom([path.join(configDir, 'config.json')]);

  const resolved = resolveConfigDir({
    xdgConfigHome: '/tmp/missing-xdg',
    homeDir,
    existsSync,
  });

  assert.equal(resolved, configDir);
});

test('resolveConfigDir falls back to existing directory when file is missing', () => {
  const homeDir = '/home/tester';
  const configDir = path.join(homeDir, '.config', 'subminer');
  const existsSync = existsSyncFrom([configDir]);

  const resolved = resolveConfigDir({
    xdgConfigHome: '/tmp/missing-xdg',
    homeDir,
    existsSync,
  });

  assert.equal(resolved, configDir);
});

test('resolveConfigFilePath prefers jsonc before json', () => {
  const homeDir = '/home/tester';
  const xdgConfigHome = '/tmp/xdg-config';
  const existsSync = existsSyncFrom([
    path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'),
    path.join(xdgConfigHome, 'SubMiner', 'config.json'),
  ]);

  const resolved = resolveConfigFilePath({
    xdgConfigHome,
    homeDir,
    existsSync,
  });

  assert.equal(resolved, path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'));
});

test('resolveConfigFilePath keeps legacy fallback output path', () => {
  const homeDir = '/home/tester';
  const xdgConfigHome = '/tmp/xdg-config';
  const existsSync = existsSyncFrom([]);

  const resolved = resolveConfigFilePath({
    xdgConfigHome,
    homeDir,
    existsSync,
  });

  assert.equal(resolved, path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'));
});
