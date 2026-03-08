import assert from 'node:assert/strict';
import test from 'node:test';
import {
  normalizeStartupArgv,
  sanitizeHelpEnv,
  sanitizeStartupEnv,
  sanitizeBackgroundEnv,
  shouldDetachBackgroundLaunch,
  shouldHandleHelpOnlyAtEntry,
} from './main-entry-runtime';

test('normalizeStartupArgv defaults no-arg startup to --start --background', () => {
  assert.deepEqual(normalizeStartupArgv(['SubMiner.AppImage'], {}), [
    'SubMiner.AppImage',
    '--start',
    '--background',
  ]);
  assert.deepEqual(
    normalizeStartupArgv(['SubMiner.AppImage', '--password-store', 'gnome-libsecret'], {}),
    ['SubMiner.AppImage', '--password-store', 'gnome-libsecret', '--start', '--background'],
  );
  assert.deepEqual(normalizeStartupArgv(['SubMiner.AppImage', '--background'], {}), [
    'SubMiner.AppImage',
    '--background',
    '--start',
  ]);
  assert.deepEqual(normalizeStartupArgv(['SubMiner.AppImage', '--help'], {}), [
    'SubMiner.AppImage',
    '--help',
  ]);
});

test('shouldHandleHelpOnlyAtEntry detects help-only invocation', () => {
  assert.equal(shouldHandleHelpOnlyAtEntry(['--help'], {}), true);
  assert.equal(shouldHandleHelpOnlyAtEntry(['--help', '--start'], {}), false);
  assert.equal(shouldHandleHelpOnlyAtEntry(['--start'], {}), false);
  assert.equal(shouldHandleHelpOnlyAtEntry(['--help'], { ELECTRON_RUN_AS_NODE: '1' }), false);
});

test('sanitizeStartupEnv suppresses warnings and lsfg layer', () => {
  const env = sanitizeStartupEnv({
    VK_INSTANCE_LAYERS: 'foo:lsfg-vk:bar',
  });
  assert.equal(env.NODE_NO_WARNINGS, '1');
  assert.equal('VK_INSTANCE_LAYERS' in env, false);
});

test('sanitizeHelpEnv suppresses warnings and lsfg layer', () => {
  const env = sanitizeHelpEnv({
    VK_INSTANCE_LAYERS: 'foo:lsfg-vk:bar',
  });
  assert.equal(env.NODE_NO_WARNINGS, '1');
  assert.equal('VK_INSTANCE_LAYERS' in env, false);
});

test('sanitizeBackgroundEnv marks background child and keeps warning suppression', () => {
  const env = sanitizeBackgroundEnv({
    VK_INSTANCE_LAYERS: 'foo:lsfg-vk:bar',
  });
  assert.equal(env.SUBMINER_BACKGROUND_CHILD, '1');
  assert.equal(env.NODE_NO_WARNINGS, '1');
  assert.equal('VK_INSTANCE_LAYERS' in env, false);
});

test('shouldDetachBackgroundLaunch only for first background invocation', () => {
  assert.equal(shouldDetachBackgroundLaunch(['--background'], {}), true);
  assert.equal(
    shouldDetachBackgroundLaunch(['--background'], { SUBMINER_BACKGROUND_CHILD: '1' }),
    false,
  );
  assert.equal(
    shouldDetachBackgroundLaunch(['--background'], { ELECTRON_RUN_AS_NODE: '1' }),
    false,
  );
  assert.equal(shouldDetachBackgroundLaunch(['--start'], {}), false);
});
