import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import net from 'node:net';
import { describe, test } from 'node:test';
import { buildMediaTimingPreviewArgs, MediaTimingPreviewSession } from './media-timing-preview';

describe('buildMediaTimingPreviewArgs', () => {
  test('creates a hidden audio-only reusable mpv session', () => {
    const args = buildMediaTimingPreviewArgs('/tmp/review.sock', {
      mediaPath: '/video/show.mkv',
      audioTrackId: 3,
      volume: 55,
    });

    assert.ok(args.includes('--no-video'));
    assert.ok(args.includes('--force-window=no'));
    assert.ok(args.includes('--idle=yes'));
    assert.ok(args.includes('--pause=yes'));
    assert.ok(args.includes('--input-ipc-server=/tmp/review.sock'));
    assert.ok(args.includes('--aid=3'));
    assert.ok(args.includes('--volume=55'));
    assert.equal(args.at(-2), '--');
    assert.equal(args.at(-1), '/video/show.mkv');
  });

  test('separates an option-like media path without adding optional audio arguments', () => {
    const args = buildMediaTimingPreviewArgs('/tmp/review.sock', {
      mediaPath: '--fullscreen',
    });

    assert.equal(args.at(-2), '--');
    assert.equal(args.at(-1), '--fullscreen');
    assert.equal(
      args.some((arg) => arg.startsWith('--aid=')),
      false,
    );
    assert.equal(
      args.some((arg) => arg.startsWith('--volume=')),
      false,
    );
  });
});

test('preview session handles socket errors after connecting', async () => {
  const socket = new net.Socket();
  const child = new EventEmitter() as EventEmitter & { kill: () => boolean };
  child.kill = () => true;
  const session = new MediaTimingPreviewSession({
    platform: 'linux',
    spawnProcess: () => child as never,
    connectSocket: () => {
      queueMicrotask(() => socket.emit('connect'));
      return socket;
    },
    removeSocketFile: () => undefined,
    createSocketPath: () => '/tmp/review.sock',
  });

  await session.start({ mediaPath: '/video/show.mkv' });
  assert.doesNotThrow(() => socket.emit('error', new Error('pipe closed')));
  await assert.rejects(session.play(1, 2), /not ready/);
  session.dispose();
});

test('preview session keeps failed connection errors handled through destruction', async () => {
  const socket = new EventEmitter() as EventEmitter & {
    destroy: () => void;
  };
  socket.destroy = () => {
    socket.emit('error', new Error('socket failed again while closing'));
  };
  const child = new EventEmitter() as EventEmitter & { kill: () => boolean };
  child.kill = () => true;
  const times = [0, 0, 0, 6_000];
  const session = new MediaTimingPreviewSession({
    platform: 'linux',
    spawnProcess: () => child as never,
    connectSocket: () => {
      queueMicrotask(() => socket.emit('error', new Error('connection failed')));
      return socket as never;
    },
    now: () => times.shift() ?? 6_000,
    removeSocketFile: () => undefined,
    createSocketPath: () => '/tmp/review.sock',
  });

  await assert.rejects(session.start({ mediaPath: '/video/show.mkv' }), /Timed out starting/);
});

test('preview session rejects a connection that finishes after disposal', async () => {
  const socket = new net.Socket();
  const child = new EventEmitter() as EventEmitter & { kill: () => boolean };
  child.kill = () => true;
  const session = new MediaTimingPreviewSession({
    platform: 'linux',
    spawnProcess: () => child as never,
    connectSocket: () => socket,
    removeSocketFile: () => undefined,
    createSocketPath: () => '/tmp/review.sock',
  });

  const pendingStart = session.start({ mediaPath: '-playlist' });
  session.dispose();
  socket.emit('connect');

  await assert.rejects(pendingStart, /closed/);
  assert.equal(socket.destroyed, true);
});

test('preview session shares one startup across concurrent start calls', async () => {
  const socket = new net.Socket();
  const child = new EventEmitter() as EventEmitter & { kill: () => boolean };
  child.kill = () => true;
  let spawnCount = 0;
  const session = new MediaTimingPreviewSession({
    platform: 'linux',
    spawnProcess: () => {
      spawnCount += 1;
      return child as never;
    },
    connectSocket: () => socket,
    removeSocketFile: () => undefined,
    createSocketPath: () => '/tmp/review.sock',
  });

  const firstStart = session.start({ mediaPath: '/video/show.mkv' });
  const secondStart = session.start({ mediaPath: '/video/show.mkv' });
  socket.emit('connect');

  await Promise.all([firstStart, secondStart]);
  assert.equal(spawnCount, 1);
  session.dispose();
});

test('preview session can start again after a startup failure', async () => {
  const socket = new net.Socket();
  const child = new EventEmitter() as EventEmitter & { kill: () => boolean };
  child.kill = () => true;
  let spawnCount = 0;
  const session = new MediaTimingPreviewSession({
    platform: 'linux',
    spawnProcess: () => {
      spawnCount += 1;
      if (spawnCount === 1) throw new Error('spawn failed');
      return child as never;
    },
    connectSocket: () => {
      queueMicrotask(() => socket.emit('connect'));
      return socket;
    },
    removeSocketFile: () => undefined,
    createSocketPath: () => '/tmp/review.sock',
  });

  await assert.rejects(session.start({ mediaPath: '/video/show.mkv' }), /spawn failed/);
  await session.start({ mediaPath: '/video/show.mkv' });
  assert.equal(spawnCount, 2);
  session.dispose();
});

test('preview session bounds a connection attempt that never settles', async () => {
  const child = new EventEmitter() as EventEmitter & { kill: () => boolean };
  child.kill = () => true;
  let nowMs = 0;
  let connectAttempts = 0;
  const session = new MediaTimingPreviewSession({
    platform: 'linux',
    spawnProcess: () => child as never,
    connectSocket: () => {
      connectAttempts += 1;
      const socket = new net.Socket();
      socket.destroy = (() => {
        socket.emit('error', new Error('socket failed while timing out'));
        return socket;
      }) as typeof socket.destroy;
      return socket;
    },
    now: () => {
      const current = nowMs;
      nowMs += 1_000;
      return current;
    },
    schedule: (callback) => setTimeout(callback, 0),
    cancelSchedule: (timeout) => clearTimeout(timeout),
    removeSocketFile: () => undefined,
    createSocketPath: () => '/tmp/review.sock',
  });

  await assert.rejects(session.start({ mediaPath: '/video/show.mkv' }), /Timed out starting/);
  assert.equal(connectAttempts, 1);
});
