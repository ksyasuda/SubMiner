import assert from 'node:assert/strict';
import test from 'node:test';
import { createNotifySendReplacer, resolveDefaultNotificationIconPath } from './notification';

test('default notification icon resolves packaged SubMiner asset when no per-notification icon is provided', () => {
  const path = resolveDefaultNotificationIconPath({
    platform: 'linux',
    resourcesPath: '/opt/SubMiner/resources',
    appPath: '/opt/SubMiner/resources/app.asar',
    dirname: '/opt/SubMiner/resources/app.asar/dist/core/utils',
    cwd: '/opt/SubMiner',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) => candidate === '/opt/SubMiner/resources/assets/SubMiner.png',
  });

  assert.equal(path, '/opt/SubMiner/resources/assets/SubMiner.png');
});

test('default notification icon prefers the square app icon when bundled images are available', () => {
  const path = resolveDefaultNotificationIconPath({
    platform: 'linux',
    resourcesPath: '/opt/SubMiner/resources',
    appPath: '/opt/SubMiner/resources/app.asar',
    dirname: '/opt/SubMiner/resources/app.asar/dist/core/utils',
    cwd: '/opt/SubMiner',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) =>
      candidate === '/opt/SubMiner/resources/assets/SubMiner.png' ||
      candidate === '/opt/SubMiner/resources/assets/SubMiner-square.png',
  });

  assert.equal(path, '/opt/SubMiner/resources/assets/SubMiner-square.png');
});

test('default notification icon avoids macOS tray template assets', () => {
  const seen: string[] = [];
  const path = resolveDefaultNotificationIconPath({
    platform: 'darwin',
    resourcesPath: '/Applications/SubMiner.app/Contents/Resources',
    appPath: '/Applications/SubMiner.app/Contents/Resources/app.asar',
    dirname: '/Applications/SubMiner.app/Contents/Resources/app.asar/dist/core/utils',
    cwd: '/Applications/SubMiner.app/Contents/Resources',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) => {
      seen.push(candidate);
      return candidate.endsWith('/assets/SubMiner-square.png');
    },
  });

  assert.equal(path, '/Applications/SubMiner.app/Contents/Resources/assets/SubMiner-square.png');
  assert.equal(
    seen.some((candidate) => candidate.includes('SubMinerTemplate')),
    false,
  );
});

test('notify-send replacer reuses the daemon-assigned id for follow-up updates', () => {
  const calls: string[][] = [];
  const replacer = createNotifySendReplacer((args, callback) => {
    calls.push(args);
    callback(null, '42\n');
  });

  replacer('sync', { title: 'SubMiner', body: 'Generating 1/3' }, () => assert.fail('no fallback'));
  replacer('sync', { title: 'SubMiner', body: 'Generating 2/3' }, () => assert.fail('no fallback'));

  assert.equal(calls.length, 2);
  assert.equal(calls[0]?.includes('--print-id'), true);
  assert.equal(
    calls[0]?.some((arg) => arg.startsWith('--replace-id=')),
    false,
  );
  assert.equal(calls[1]?.includes('--replace-id=42'), true);
});

test('notify-send replacer collapses a mid-send burst to the latest update', () => {
  const pending: Array<(error: Error | null, stdout: string) => void> = [];
  const calls: string[][] = [];
  const replacer = createNotifySendReplacer((args, callback) => {
    calls.push(args);
    pending.push(callback);
  });

  replacer('sync', { title: 'SubMiner', body: 'first' }, () => assert.fail('no fallback'));
  replacer('sync', { title: 'SubMiner', body: 'second' }, () => assert.fail('no fallback'));
  replacer('sync', { title: 'SubMiner', body: 'third' }, () => assert.fail('no fallback'));

  // Only one send is in flight, so the id capture cannot race.
  assert.equal(calls.length, 1);
  assert.equal(calls[0]?.at(-1), 'first');

  pending[0]?.(null, '7');

  // 'second' is stale by the time the daemon frees up, so only 'third' is sent.
  assert.equal(calls.length, 2);
  assert.equal(calls[1]?.at(-1), 'third');
  assert.equal(calls[1]?.includes('--replace-id=7'), true);

  pending[1]?.(null, '7');
  assert.equal(calls.length, 2);
});

test('notify-send replacer escapes freedesktop body markup', () => {
  const calls: string[][] = [];
  const replacer = createNotifySendReplacer((args, callback) => {
    calls.push(args);
    callback(null, '1');
  });

  replacer('sync', { title: 'SubMiner', body: 'Steins;Gate <0 & more' }, () => {});

  assert.equal(calls[0]?.at(-1), 'Steins;Gate &lt;0 &amp; more');
});

test('notify-send replacer falls back to Electron notifications permanently after a failure', () => {
  let execCalls = 0;
  let fallbacks = 0;
  const replacer = createNotifySendReplacer((_args, callback) => {
    execCalls += 1;
    callback(new Error('spawn notify-send ENOENT'), '');
  });

  replacer('sync', { title: 'SubMiner', body: 'first' }, () => (fallbacks += 1));
  replacer('sync', { title: 'SubMiner', body: 'second' }, () => (fallbacks += 1));

  assert.equal(execCalls, 1);
  assert.equal(fallbacks, 2);
});

test('notify-send replacer falls back for an update queued before the failure lands', () => {
  let execCalls = 0;
  let fallbacks = 0;
  const pending: Array<(error: Error | null, stdout: string) => void> = [];
  const replacer = createNotifySendReplacer((_args, callback) => {
    execCalls += 1;
    pending.push(callback);
  });

  replacer('sync', { title: 'SubMiner', body: 'first' }, () => (fallbacks += 1));
  replacer('sync', { title: 'SubMiner', body: 'second' }, () => (fallbacks += 1));
  // The coalesced update is still waiting when the in-flight send fails, so it must fall back too
  // instead of being dropped.
  pending[0]?.(new Error('spawn notify-send ENOENT'), '');

  assert.equal(execCalls, 1);
  assert.equal(fallbacks, 2);
});

test('notify-send replacer survives a synchronous spawn throw', () => {
  let execCalls = 0;
  let fallbacks = 0;
  const replacer = createNotifySendReplacer(() => {
    execCalls += 1;
    throw new Error('spawn threw synchronously');
  });

  replacer('sync', { title: 'SubMiner', body: 'first' }, () => (fallbacks += 1));
  // A wedged entry would silently drop every later update, so the fallback must still fire.
  replacer('sync', { title: 'SubMiner', body: 'second' }, () => (fallbacks += 1));

  assert.equal(execCalls, 1);
  assert.equal(fallbacks, 2);
});

test('default notification icon resolves cwd fallback through injected deps', () => {
  const resolvedPath = resolveDefaultNotificationIconPath({
    platform: 'linux',
    resourcesPath: '/missing/resources',
    appPath: '/missing/app',
    dirname: '/missing/dist/core/utils',
    cwd: '/portable/SubMiner',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) => candidate === '/portable/SubMiner/assets/SubMiner-square.png',
  });

  assert.equal(resolvedPath, '/portable/SubMiner/assets/SubMiner-square.png');
});
