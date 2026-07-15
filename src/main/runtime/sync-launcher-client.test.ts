import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import type { SyncProgressEvent } from '../../shared/sync/sync-events';
import {
  runSyncLauncher,
  sanitizeSyncLauncherEnv,
  type SyncLauncherSpawn,
} from './sync-launcher-client';

class FakeChild extends EventEmitter {
  stdout = new EventEmitter();
  stderr = new EventEmitter();
  killed = false;
  kill(): boolean {
    this.killed = true;
    this.emit('close', null, 'SIGTERM');
    return true;
  }
}

function makeSpawn(): { spawn: SyncLauncherSpawn; children: FakeChild[]; commands: string[][] } {
  const children: FakeChild[] = [];
  const commands: string[][] = [];
  const spawn: SyncLauncherSpawn = (command, args) => {
    commands.push([command, ...args]);
    const child = new FakeChild();
    children.push(child);
    return child as never;
  };
  return { spawn, children, commands };
}

test('runSyncLauncher parses NDJSON events across chunk boundaries', async () => {
  const { spawn, children, commands } = makeSpawn();
  const events: SyncProgressEvent[] = [];
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--json'],
    onEvent: (event) => events.push(event),
    spawn,
  });

  const child = children[0]!;
  child.stdout.emit('data', Buffer.from('{"type":"stage","stage":"snapshot-'));
  child.stdout.emit('data', Buffer.from('local","message":"Snapshotting"}\n{"type":"result",'));
  child.stdout.emit('data', Buffer.from('"ok":true,"error":null}\n'));
  child.emit('close', 0, null);

  const result = await handle.done;
  assert.deepEqual(commands[0], ['subminer', 'sync', 'media-box', '--json']);
  assert.equal(events.length, 2);
  assert.equal(events[0]!.type, 'stage');
  assert.deepEqual(events[1], { type: 'result', ok: true, error: null });
  assert.deepEqual(result, { ok: true, error: null });
});

test('runSyncLauncher preserves multibyte characters split across chunks', async () => {
  const { spawn, children } = makeSpawn();
  const events: SyncProgressEvent[] = [];
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--json'],
    onEvent: (event) => events.push(event),
    spawn,
  });

  const child = children[0]!;
  const line = Buffer.from('{"type":"result","ok":false,"error":"進捗の同期に失敗"}\n');
  // Split mid-character: the boundary lands inside a multibyte code point.
  const boundary = line.indexOf(Buffer.from('進')) + 1;
  child.stdout.emit('data', line.subarray(0, boundary));
  child.stdout.emit('data', line.subarray(boundary));
  child.emit('close', 1, null);

  const result = await handle.done;
  assert.deepEqual(events.at(-1), {
    type: 'result',
    ok: false,
    error: '進捗の同期に失敗',
  });
  assert.equal(result.error, '進捗の同期に失敗');
});

test('runSyncLauncher settles after exit when close never arrives', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--json'],
    onEvent: () => {},
    spawn,
  });
  const child = children[0]!;
  child.stderr.emit('data', Buffer.from('[ERROR] remote refused\n'));
  // The child is gone, but a descendant still holds the inherited stdio pipes,
  // so `close` never fires.
  child.emit('exit', 1, null);

  const result = await handle.done;
  assert.equal(result.ok, false);
  assert.match(result.error ?? '', /remote refused/);
});

test('runSyncLauncher reports failures with the result event error or stderr tail', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--json'],
    onEvent: () => {},
    spawn,
  });
  const child = children[0]!;
  child.stdout.emit(
    'data',
    Buffer.from('{"type":"result","ok":false,"error":"Remote merge failed"}\n'),
  );
  child.stderr.emit('data', Buffer.from('[ERROR] Remote merge failed\n'));
  child.emit('close', 1, null);

  const result = await handle.done;
  assert.equal(result.ok, false);
  assert.match(result.error ?? '', /Remote merge failed/);
});

test('runSyncLauncher falls back to stderr when no result event arrives', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', '--snapshot', '/tmp/x.sqlite', '--json'],
    onEvent: () => {},
    spawn,
  });
  const child = children[0]!;
  child.stderr.emit('data', Buffer.from('[ERROR] boom\n'));
  child.emit('close', 1, null);

  const result = await handle.done;
  assert.equal(result.ok, false);
  assert.match(result.error ?? '', /boom/);
});

test('runSyncLauncher settles on exit when the terminal event is already parsed', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--check', '--json'],
    onEvent: () => {},
    spawn,
  });
  children[0]!.stdout.emit('data', Buffer.from('{"type":"result","ok":true,"error":null}\n'));
  children[0]!.emit('exit', 0, null);

  const result = await Promise.race([
    handle.done,
    new Promise<null>((resolve) => setTimeout(() => resolve(null), 25)),
  ]);
  assert.deepEqual(result, { ok: true, error: null });
});

test('runSyncLauncher waits for terminal NDJSON when exit precedes final stdout data', async () => {
  const { spawn, children } = makeSpawn();
  const events: SyncProgressEvent[] = [];
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--json'],
    onEvent: (event) => events.push(event),
    spawn,
  });
  const child = children[0]!;
  let settled = false;
  void handle.done.then(() => {
    settled = true;
  });

  child.emit('exit', 0, null);
  await Promise.resolve();
  assert.equal(settled, false);

  child.stdout.emit('data', Buffer.from('{"type":"result","ok":true,"error":null}\n'));

  assert.deepEqual(await handle.done, { ok: true, error: null });
  assert.deepEqual(events.at(-1), { type: 'result', ok: true, error: null });
});

test('runSyncLauncher cancel kills the child and resolves as cancelled', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--json'],
    onEvent: () => {},
    spawn,
  });
  const child = children[0]!;
  child.kill = () => {
    child.killed = true;
    child.emit('exit', null, 'SIGTERM');
    return true;
  };
  handle.cancel();
  assert.equal(child.killed, true);

  const result = await Promise.race([
    handle.done,
    new Promise<null>((resolve) => setTimeout(() => resolve(null), 25)),
  ]);
  assert.ok(result);
  assert.equal(result.ok, false);
  assert.match(result.error ?? '', /cancel/i);
});

test('runSyncLauncher cancel settles a child that already exited without a result', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--json'],
    onEvent: () => {},
    spawn,
  });
  const child = children[0]!;
  // The child is gone, but a descendant still holds the inherited stdio pipes,
  // so `close` never arrives.
  child.kill = () => {
    child.killed = true;
    return true;
  };
  child.emit('exit', null, 'SIGTERM');

  handle.cancel();

  const result = await Promise.race([
    handle.done,
    new Promise<null>((resolve) => setTimeout(() => resolve(null), 25)),
  ]);
  assert.deepEqual(result, { ok: false, error: 'Sync cancelled.' });
  assert.equal(child.killed, false);
});

test('runSyncLauncher times out a child that already exited without a result', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--check', '--json'],
    onEvent: () => {},
    spawn,
    timeoutMs: 1,
  });
  const child = children[0]!;
  child.kill = () => {
    child.killed = true;
    return true;
  };
  child.emit('exit', null, 'SIGTERM');

  const result = await Promise.race([
    handle.done,
    new Promise<null>((resolve) => setTimeout(() => resolve(null), 25)),
  ]);
  assert.deepEqual(result, { ok: false, error: 'Sync operation timed out.' });
  assert.equal(child.killed, false);
});

test('runSyncLauncher times out a child that never completes', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--check', '--json'],
    onEvent: () => {},
    spawn,
    timeoutMs: 1,
  });
  const child = children[0]!;
  child.kill = () => {
    child.killed = true;
    child.emit('exit', null, 'SIGTERM');
    return true;
  };

  const result = await Promise.race([
    handle.done,
    new Promise<null>((resolve) => setTimeout(() => resolve(null), 25)),
  ]);
  assert.equal(child.killed, true);
  assert.deepEqual(result, { ok: false, error: 'Sync operation timed out.' });
});

test('runSyncLauncher surfaces spawn errors', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--json'],
    onEvent: () => {},
    spawn,
  });
  children[0]!.emit('error', new Error('ENOENT: subminer not found'));

  const result = await handle.done;
  assert.equal(result.ok, false);
  assert.match(result.error ?? '', /ENOENT/);
});

test('resolveSyncLauncherCommand self-spawns the app in --sync-cli mode', async () => {
  const { resolveSyncLauncherCommand } = await import('./sync-launcher-client');
  const packaged = resolveSyncLauncherCommand({
    execPath: '/opt/SubMiner/subminer-app',
    appPath: null,
  });
  assert.deepEqual(packaged, ['/opt/SubMiner/subminer-app', '--sync-cli']);

  const dev = resolveSyncLauncherCommand({
    execPath: '/usr/bin/electron',
    appPath: '/home/u/SubMiner',
  });
  assert.deepEqual(dev, ['/usr/bin/electron', '/home/u/SubMiner', '--sync-cli']);
});

test('sanitizeSyncLauncherEnv removes transported parent startup arguments', () => {
  const parentEnv = {
    PATH: '/usr/bin',
    ELECTRON_RUN_AS_NODE: '1',
    SUBMINER_APP_ARGC: '2',
    SUBMINER_APP_ARG_0: '--start',
    SUBMINER_APP_ARG_1: '--background',
  };

  assert.deepEqual(sanitizeSyncLauncherEnv(parentEnv), { PATH: '/usr/bin' });
  assert.equal(parentEnv.SUBMINER_APP_ARGC, '2');
});
