import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { EventEmitter } from 'node:events';
import type { spawn as spawnType, ChildProcess } from 'node:child_process';
import { allocatePort, startSidecar } from './sidecar-process';
import type { BundleBinaries } from './sidecar-bundle';

const binaries: BundleBinaries = {
  javaPath: '/nonexistent/java',
  jarPath: '/tmp/MExtensionServer.jar',
};

/** A ChildProcess stand-in: an EventEmitter with the bits startSidecar touches. */
function fakeChild(onKill?: (child: EventEmitter) => void): ChildProcess {
  const child = new EventEmitter();
  Object.assign(child, {
    stdout: null,
    stderr: null,
    kill: () => {
      onKill?.(child);
      return true;
    },
  });
  return child as unknown as ChildProcess;
}

test('a failed spawn rejects instead of throwing an unhandled error event', async () => {
  const port = await allocatePort();
  const child = fakeChild();
  const spawnImpl = (() => {
    // Node emits `error` asynchronously when the binary cannot be executed.
    queueMicrotask(() => child.emit('error', new Error('spawn ENOENT')));
    return child;
  }) as unknown as typeof spawnType;

  await assert.rejects(
    () => startSidecar({ binaries, port, readyTimeoutMs: 2000, spawnImpl }),
    /could not start.*ENOENT/,
  );
});

test('a readiness timeout shuts the child down and reports the timeout', async () => {
  const port = await allocatePort();
  // Never becomes ready, but does go down on the first signal.
  const child = fakeChild((emitter) => queueMicrotask(() => emitter.emit('exit', 0, 'SIGTERM')));
  const spawnImpl = (() => child) as unknown as typeof spawnType;

  await assert.rejects(
    () => startSidecar({ binaries, port, readyTimeoutMs: 50, spawnImpl }),
    /did not become ready within 50ms/,
  );
});

test('an early exit is reported with its code rather than waiting out the deadline', async () => {
  const port = await allocatePort();
  const child = fakeChild();
  const spawnImpl = (() => {
    queueMicrotask(() => child.emit('exit', 1, null));
    return child;
  }) as unknown as typeof spawnType;

  await assert.rejects(
    () => startSidecar({ binaries, port, readyTimeoutMs: 2000, spawnImpl }),
    /exited before becoming ready \(code 1/,
  );
});

test('onExit reports a death after readiness, including to late subscribers', async () => {
  const port = await allocatePort();
  // Fake the bridge's capabilities endpoint so startSidecar reports ready.
  const server = http.createServer((_req, res) => {
    res.writeHead(200, { 'content-type': 'application/json' });
    res.end(
      JSON.stringify({ mangatanMihonBridge: 1, sourceFactory: true, preferenceCallbacks: true }),
    );
  });
  await new Promise<void>((resolve) => server.listen(port, '127.0.0.1', resolve));
  const child = fakeChild();
  const spawnImpl = (() => child) as unknown as typeof spawnType;

  try {
    const handle = await startSidecar({ binaries, port, readyTimeoutMs: 5000, spawnImpl });
    const exits: Array<{ code: number | null; signal: NodeJS.Signals | null }> = [];
    handle.onExit((info) => exits.push(info));
    assert.equal(exits.length, 0);

    child.emit('exit', 137, null);
    await new Promise((resolve) => setImmediate(resolve));
    assert.deepEqual([...exits], [{ code: 137, signal: null }]);

    // A listener attached after the death still hears about it.
    handle.onExit((info) => exits.push(info));
    await new Promise((resolve) => setImmediate(resolve));
    assert.deepEqual(
      [...exits],
      [
        { code: 137, signal: null },
        { code: 137, signal: null },
      ],
    );
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
});
