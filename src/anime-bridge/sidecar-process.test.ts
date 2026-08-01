import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import type { spawn as spawnType, ChildProcess } from 'node:child_process';
import { allocatePort, startSidecar } from './sidecar-process';
import type { BundleBinaries } from './sidecar-bundle';

const binaries: BundleBinaries = {
  javaPath: '/nonexistent/java',
  jarPath: '/tmp/MExtensionServer.jar',
};

/** A ChildProcess stand-in: an EventEmitter with the bits startSidecar touches. */
function fakeChild(): ChildProcess {
  const child = new EventEmitter();
  Object.assign(child, { stdout: null, stderr: null, kill: () => true });
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
