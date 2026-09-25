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

  test('keeps source timestamps for cached remote windows', () => {
    const args = buildMediaTimingPreviewArgs('/tmp/review.sock', {
      mediaPath: '/tmp/window.mkv',
      absoluteTimestamps: true,
    });

    assert.ok(args.includes('--rebase-start-time=no'));
    assert.equal(
      buildMediaTimingPreviewArgs('/tmp/review.sock', { mediaPath: '/video/show.mkv' }).includes(
        '--rebase-start-time=no',
      ),
      false,
    );
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

function createFakeSocket() {
  const socket = new EventEmitter() as EventEmitter & {
    destroyed: boolean;
    write: (data: string) => boolean;
    end: () => void;
    destroy: () => void;
    off: EventEmitter['off'];
  };
  const writes: string[] = [];
  socket.destroyed = false;
  socket.write = (data) => {
    writes.push(data);
    return true;
  };
  socket.end = () => undefined;
  socket.destroy = () => {
    socket.destroyed = true;
  };
  return { socket, writes };
}

test('preview session plays once to the clip end and reports when mpv has drained it', async () => {
  const { socket, writes } = createFakeSocket();
  const child = new EventEmitter() as EventEmitter & { kill: () => boolean };
  child.kill = () => true;
  const session = new MediaTimingPreviewSession({
    platform: 'linux',
    spawnProcess: () => child as never,
    connectSocket: () => {
      queueMicrotask(() => socket.emit('connect'));
      return socket as never;
    },
    removeSocketFile: () => undefined,
    createSocketPath: () => '/tmp/review.sock',
  });
  let endedCount = 0;
  session.onPlaybackEnded(() => {
    endedCount += 1;
  });
  const property = (name: string, data: boolean): string =>
    `${JSON.stringify({ event: 'property-change', name, data })}\n`;

  await session.start({ mediaPath: '/video/show.mkv' });
  assert.deepEqual(
    writes.map((line) => JSON.parse(line).command),
    [
      ['observe_property', 1, 'eof-reached'],
      ['observe_property', 2, 'pause'],
    ],
  );
  // The observers' initial replies describe the idle paused player, not a finished preview.
  socket.emit('data', property('eof-reached', false) + property('pause', true));
  assert.equal(endedCount, 0);

  writes.length = 0;
  await session.play(12.25, 14.5);
  assert.deepEqual(
    writes.map((line) => JSON.parse(line).command),
    [
      ['set_property', 'pause', true],
      ['seek', 12.25, 'absolute+exact'],
      ['set_property', 'end', '14.500'],
      ['set_property', 'pause', false],
    ],
  );

  // Events may arrive split across chunks. The decoder passing `end` flips eof-reached while
  // audio still drains; only the keep-open pause that follows marks the preview as finished.
  socket.emit('data', property('eof-reached', false) + property('pause', false).slice(0, 20));
  socket.emit('data', property('pause', false).slice(20) + property('eof-reached', true));
  assert.equal(endedCount, 0);
  socket.emit('data', property('pause', true));
  assert.equal(endedCount, 1);
  socket.emit('data', property('pause', true));
  assert.equal(endedCount, 1);

  // Stopping early pauses without an end signal, and a later real EOF is not a preview end.
  await session.play(1, 2);
  socket.emit('data', property('eof-reached', false) + property('pause', false));
  await session.stop();
  socket.emit('data', property('pause', true) + property('eof-reached', true));
  assert.equal(endedCount, 1);
  session.dispose();
});
