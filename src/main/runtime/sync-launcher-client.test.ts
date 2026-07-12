import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import type { SyncProgressEvent } from '../../shared/sync/sync-events';
import { runSyncLauncher, type SyncLauncherSpawn } from './sync-launcher-client';

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

test('runSyncLauncher cancel kills the child and resolves as cancelled', async () => {
  const { spawn, children } = makeSpawn();
  const handle = runSyncLauncher({
    command: ['subminer'],
    args: ['sync', 'media-box', '--json'],
    onEvent: () => {},
    spawn,
  });
  handle.cancel();
  assert.equal(children[0]!.killed, true);

  const result = await handle.done;
  assert.equal(result.ok, false);
  assert.match(result.error ?? '', /cancel/i);
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

test('resolveSyncLauncherCommand prefers the bundled launcher over PATH', async () => {
  const { resolveSyncLauncherCommand } = await import('./sync-launcher-client');
  const bundled = resolveSyncLauncherCommand({
    findCommand: (name: string) =>
      name === 'bun' ? '/usr/bin/bun' : name === 'subminer' ? '/home/u/.local/bin/subminer' : null,
    resolveResourcePath: () => '/opt/app/resources/launcher/subminer',
    existsSync: () => true,
  });
  assert.deepEqual(bundled.command, ['/usr/bin/bun', '/opt/app/resources/launcher/subminer']);

  const fallback = resolveSyncLauncherCommand({
    findCommand: (name: string) => (name === 'subminer' ? '/home/u/.local/bin/subminer' : null),
    resolveResourcePath: () => '/missing',
    existsSync: () => false,
  });
  assert.deepEqual(fallback.command, ['/home/u/.local/bin/subminer']);

  const none = resolveSyncLauncherCommand({
    findCommand: () => null,
    resolveResourcePath: () => '/missing',
    existsSync: () => false,
  });
  assert.equal(none.command, null);
  assert.match(none.error ?? '', /launcher/i);
});
