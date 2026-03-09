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
  const trimmedXdgConfigHome = '/home/tester/.config';
  const fallbackDir = path.join(homeDir, '.config');
  const baseDirs = resolveConfigBaseDirs(` ${trimmedXdgConfigHome} `, homeDir, 'linux');
  const expected = Array.from(new Set([trimmedXdgConfigHome, fallbackDir]));
  assert.deepEqual(baseDirs, expected);
});

test('resolveConfigBaseDirs prefers APPDATA on windows and deduplicates fallback dir', () => {
  const homeDir = 'C:\\Users\\tester';
  const appDataDir = 'C:\\Users\\tester\\AppData\\Roaming';

  const baseDirs = resolveConfigBaseDirs(undefined, homeDir, 'win32', ` ${appDataDir} `);

  assert.deepEqual(baseDirs, [appDataDir]);
});

test('resolveConfigDir prefers xdg SubMiner config when present', () => {
  const homeDir = '/home/tester';
  const xdgConfigHome = '/tmp/xdg-config';
  const configDir = path.join(xdgConfigHome, 'SubMiner');
  const existsSync = existsSyncFrom([path.join(configDir, 'config.jsonc')]);

  const resolved = resolveConfigDir({
    xdgConfigHome,
    homeDir,
    platform: 'linux',
    existsSync,
  });

  assert.equal(resolved, configDir);
});

test('resolveConfigDir ignores lowercase subminer candidate', () => {
  const homeDir = '/home/tester';
  const lowercaseConfigDir = path.join(homeDir, '.config', 'subminer');
  const existsSync = existsSyncFrom([path.join(lowercaseConfigDir, 'config.json')]);

  const resolved = resolveConfigDir({
    xdgConfigHome: '/tmp/missing-xdg',
    homeDir,
    platform: 'linux',
    existsSync,
  });

  assert.equal(resolved, path.join('/tmp/missing-xdg', 'SubMiner'));
});

test('resolveConfigDir falls back to existing directory when file is missing', () => {
  const homeDir = '/home/tester';
  const configDir = path.join(homeDir, '.config', 'SubMiner');
  const existsSync = existsSyncFrom([configDir]);

  const resolved = resolveConfigDir({
    xdgConfigHome: '/tmp/missing-xdg',
    homeDir,
    platform: 'linux',
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
    platform: 'linux',
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
    platform: 'linux',
    existsSync,
  });

  assert.equal(resolved, path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'));
});

test('resolveConfigDir prefers APPDATA SubMiner config on windows when present', () => {
  const homeDir = 'C:\\Users\\tester';
  const appDataDir = 'C:\\Users\\tester\\AppData\\Roaming';
  const configDir = path.join(appDataDir, 'SubMiner');
  const existsSync = existsSyncFrom([path.join(configDir, 'config.jsonc')]);

  const resolved = resolveConfigDir({
    platform: 'win32',
    appDataDir,
    homeDir,
    existsSync,
  });

  assert.equal(resolved, configDir);
});

test('resolveConfigFilePath uses APPDATA fallback output path on windows', () => {
  const homeDir = 'C:\\Users\\tester';
  const appDataDir = 'C:\\Users\\tester\\AppData\\Roaming';
  const existsSync = existsSyncFrom([]);

  const resolved = resolveConfigFilePath({
    platform: 'win32',
    appDataDir,
    homeDir,
    existsSync,
  });

  assert.equal(resolved, path.join(appDataDir, 'SubMiner', 'config.jsonc'));
});
