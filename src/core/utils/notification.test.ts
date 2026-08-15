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

test('notify-send replacer passes the resolved icon path through', () => {
  const calls: string[][] = [];
  const replacer = createNotifySendReplacer((args, callback) => {
    calls.push(args);
    callback(null, '3');
  });

  replacer(
    'sync',
    { title: 'SubMiner', body: 'Generating', iconPath: '/opt/SubMiner/assets/SubMiner-square.png' },
    () => assert.fail('no fallback'),
  );

  assert.equal(calls[0]?.includes('--icon=/opt/SubMiner/assets/SubMiner-square.png'), true);
});

test('notify-send replacer tracks a separate daemon id per replaceId', () => {
  const calls: string[][] = [];
  let nextId = 10;
  const replacer = createNotifySendReplacer((args, callback) => {
    calls.push(args);
    callback(null, String(nextId++));
  });

  replacer('dictionary', { title: 'SubMiner', body: 'dict 1' }, () => assert.fail('no fallback'));
  replacer('startup', { title: 'SubMiner', body: 'startup 1' }, () => assert.fail('no fallback'));
  replacer('dictionary', { title: 'SubMiner', body: 'dict 2' }, () => assert.fail('no fallback'));
  replacer('startup', { title: 'SubMiner', body: 'startup 2' }, () => assert.fail('no fallback'));

  assert.equal(calls.length, 4);
  assert.equal(calls[2]?.includes('--replace-id=10'), true);
  assert.equal(calls[3]?.includes('--replace-id=11'), true);
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

function spawnError(code: string): Error {
  return Object.assign(new Error(`spawn notify-send ${code}`), { code });
}

function daemonError(): Error {
  // execFile reports a bad exit as a numeric code, unlike a libuv spawn failure.
  return Object.assign(new Error('notify-send exited with code 1'), { code: 1 });
}

test('notify-send replacer falls back to Electron notifications permanently when the binary is missing', () => {
  let execCalls = 0;
  let fallbacks = 0;
  const replacer = createNotifySendReplacer((_args, callback) => {
    execCalls += 1;
    callback(spawnError('ENOENT'), '');
  });

  replacer('sync', { title: 'SubMiner', body: 'first' }, () => (fallbacks += 1));
  replacer('sync', { title: 'SubMiner', body: 'second' }, () => (fallbacks += 1));

  assert.equal(execCalls, 1);
  assert.equal(fallbacks, 2);
});

test('notify-send replacer keeps using notify-send after a transient daemon failure', () => {
  let execCalls = 0;
  let fallbacks = 0;
  const replacer = createNotifySendReplacer((_args, callback) => {
    execCalls += 1;
    callback(execCalls === 1 ? daemonError() : null, '5');
  });

  replacer('sync', { title: 'SubMiner', body: 'first' }, () => (fallbacks += 1));
  replacer('sync', { title: 'SubMiner', body: 'second' }, () => assert.fail('no fallback'));

  // One bad answer only costs that update; a busy daemon must not downgrade the rest of the run.
  assert.equal(execCalls, 2);
  assert.equal(fallbacks, 1);
});

test('notify-send replacer retries after a transient spawn failure', () => {
  let execCalls = 0;
  let fallbacks = 0;
  const replacer = createNotifySendReplacer((_args, callback) => {
    execCalls += 1;
    // EMFILE means the process was out of descriptors, not that notify-send is unusable.
    callback(execCalls === 1 ? spawnError('EMFILE') : null, '5');
  });

  replacer('sync', { title: 'SubMiner', body: 'first' }, () => (fallbacks += 1));
  replacer('sync', { title: 'SubMiner', body: 'second' }, () => assert.fail('no fallback'));

  assert.equal(execCalls, 2);
  assert.equal(fallbacks, 1);
});

test('notify-send replacer gives up after a run of transient failures', () => {
  let execCalls = 0;
  let fallbacks = 0;
  const replacer = createNotifySendReplacer((_args, callback) => {
    execCalls += 1;
    callback(daemonError(), '');
  });

  for (let update = 0; update < 5; update += 1) {
    replacer('sync', { title: 'SubMiner', body: `update ${update}` }, () => (fallbacks += 1));
  }

  // Three strikes, then every later update goes straight to Electron instead of waiting on a spawn.
  assert.equal(execCalls, 3);
  assert.equal(fallbacks, 5);
});

test('notify-send replacer falls back with only the latest update when a send fails mid-flight', () => {
  let execCalls = 0;
  const fallbacks: string[] = [];
  const pending: Array<(error: Error | null, stdout: string) => void> = [];
  const replacer = createNotifySendReplacer((_args, callback) => {
    execCalls += 1;
    pending.push(callback);
  });

  replacer('sync', { title: 'SubMiner', body: 'first' }, () => fallbacks.push('first'));
  replacer('sync', { title: 'SubMiner', body: 'second' }, () => fallbacks.push('second'));
  pending[0]?.(spawnError('ENOENT'), '');

  assert.equal(execCalls, 1);
  // The queued update still reaches the user, and the superseded one is dropped rather than
  // flashing a stale message ahead of it.
  assert.deepEqual(fallbacks, ['second']);
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
