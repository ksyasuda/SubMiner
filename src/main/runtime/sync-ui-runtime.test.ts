import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { SyncProgressEvent } from '../../shared/sync/sync-events';
import type { SyncLauncherRunHandle } from './sync-launcher-client';
import { createSyncUiRuntime, type SyncUiRuntimeDeps } from './sync-ui-runtime';

interface LauncherCall {
  args: string[];
  timeoutMs: number | undefined;
  onEvent: (event: SyncProgressEvent) => void;
  finish: (result: { ok: boolean; error: string | null }) => void;
  cancelled: boolean;
}

function makeTestRig(root: string, overrides: Partial<SyncUiRuntimeDeps> = {}) {
  const launcherCalls: LauncherCall[] = [];
  const sent: Array<{ channel: string; payload: unknown }> = [];
  const handlers = new Map<string, (event: unknown, ...args: unknown[]) => unknown>();
  const windowStub = {
    isDestroyed: () => false,
    webContents: {
      send: (channel: string, payload?: unknown) => sent.push({ channel, payload }),
    },
  };

  const deps: SyncUiRuntimeDeps = {
    ipcMain: {
      handle: (channel: string, listener: (event: unknown, ...args: unknown[]) => unknown) => {
        handlers.set(channel, listener);
      },
    },
    hostsFilePath: path.join(root, 'sync-hosts.json'),
    snapshotsDir: path.join(root, 'snapshots'),
    getDbPath: () => path.join(root, 'immersion.sqlite'),
    resolveLauncherCommand: () => ({ command: ['subminer'], error: null }),
    runLauncher: (options): SyncLauncherRunHandle => {
      let finish: (result: { ok: boolean; error: string | null }) => void = () => {};
      const done = new Promise<{ ok: boolean; error: string | null }>((resolve) => {
        finish = resolve;
      });
      const call: LauncherCall = {
        args: options.args,
        timeoutMs: options.timeoutMs,
        onEvent: options.onEvent,
        finish,
        cancelled: false,
      };
      launcherCalls.push(call);
      return {
        cancel: () => {
          call.cancelled = true;
          finish({ ok: false, error: 'Sync cancelled.' });
        },
        done,
      };
    },
    getWindow: () => windowStub,
    pickSnapshotFile: async () => null,
    revealPath: () => {},
    nowMs: () => 1700000000000,
    ...overrides,
  };

  const runtime = createSyncUiRuntime(deps);
  runtime.registerHandlers();
  const invoke = (channel: string, ...args: unknown[]): unknown => {
    const handler = handlers.get(channel);
    assert.ok(handler, `missing handler for ${channel}`);
    return handler({}, ...args);
  };
  return { runtime, invoke, launcherCalls, sent, handlers };
}

function withTempDir(fn: (dir: string) => Promise<void> | void): Promise<void> | void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-sync-ui-test-'));
  const cleanup = () => fs.rmSync(dir, { recursive: true, force: true });
  const result = fn(dir);
  if (result instanceof Promise) return result.finally(cleanup);
  cleanup();
  return result;
}

test('registerHandlers registers every sync-ui request channel', () =>
  withTempDir((root) => {
    const { handlers } = makeTestRig(root);
    for (const channel of [
      'sync-ui:get-snapshot',
      'sync-ui:save-host',
      'sync-ui:remove-host',
      'sync-ui:set-auto-sync-interval',
      'sync-ui:run-sync',
      'sync-ui:cancel-run',
      'sync-ui:check-host',
      'sync-ui:create-snapshot',
      'sync-ui:merge-snapshot-file',
      'sync-ui:delete-snapshot',
      'sync-ui:reveal-snapshot',
      'sync-ui:pick-snapshot-file',
    ]) {
      assert.ok(handlers.has(channel), `missing ${channel}`);
    }
  }));

test('save/remove host round-trips through the hosts file', () =>
  withTempDir(async (root) => {
    const { invoke } = makeTestRig(root);
    const saved = (await invoke('sync-ui:save-host', {
      host: 'user@laptop',
      label: 'Laptop',
      direction: 'pull',
      autoSync: true,
    })) as { hosts: Array<{ host: string; direction: string; autoSync: boolean }> };
    assert.equal(saved.hosts.length, 1);
    assert.equal(saved.hosts[0]!.direction, 'pull');
    assert.equal(saved.hosts[0]!.autoSync, true);

    const removed = (await invoke('sync-ui:remove-host', 'user@laptop')) as {
      hosts: unknown[];
    };
    assert.equal(removed.hosts.length, 0);
  }));

test('save-host rejects invalid hosts', () =>
  withTempDir(async (root) => {
    const { invoke } = makeTestRig(root);
    await assert.rejects(
      async () => invoke('sync-ui:save-host', { host: '-oProxyCommand=evil' }),
      /Invalid sync host/,
    );
  }));

test('run-sync spawns the launcher, forwards progress, and rejects concurrent runs', () =>
  withTempDir(async (root) => {
    const { invoke, launcherCalls, sent } = makeTestRig(root);
    const start = (await invoke('sync-ui:run-sync', {
      host: 'media-box',
      direction: 'push',
      force: true,
    })) as { started: boolean; runId: number };
    assert.equal(start.started, true);
    assert.deepEqual(launcherCalls[0]!.args, ['sync', 'media-box', '--push', '--force', '--json']);

    const second = (await invoke('sync-ui:run-sync', { host: 'other' })) as {
      started: boolean;
      reason: string | null;
    };
    assert.equal(second.started, false);
    assert.match(second.reason ?? '', /already running/i);

    launcherCalls[0]!.onEvent({ type: 'stage', stage: 'snapshot-local', message: 'Snap' });
    const progress = sent.filter((entry) => entry.channel === 'sync-ui:progress');
    assert.equal(progress.length, 1);
    const payload = progress[0]!.payload as { runId: number; host: string; event: unknown };
    assert.equal(payload.runId, start.runId);
    assert.equal(payload.host, 'media-box');

    launcherCalls[0]!.finish({ ok: true, error: null });
    await new Promise((resolve) => setImmediate(resolve));
    assert.ok(sent.some((entry) => entry.channel === 'sync-ui:state-changed'));

    const third = (await invoke('sync-ui:run-sync', { host: 'media-box' })) as {
      started: boolean;
    };
    assert.equal(third.started, true);
  }));

test('run-sync saves the host so previously used hosts persist', () =>
  withTempDir(async (root) => {
    const { invoke, launcherCalls } = makeTestRig(root);
    await invoke('sync-ui:run-sync', { host: 'media-box', direction: 'pull' });
    const snapshot = (await invoke('sync-ui:get-snapshot')) as {
      hosts: { hosts: Array<{ host: string; direction: string }> };
    };
    assert.equal(snapshot.hosts.hosts[0]!.host, 'media-box');
    assert.equal(snapshot.hosts.hosts[0]!.direction, 'pull');
    launcherCalls[0]!.finish({ ok: true, error: null });
  }));

test('create-snapshot runs the launcher with a timestamped path under the snapshots dir', () =>
  withTempDir(async (root) => {
    const { invoke, launcherCalls } = makeTestRig(root);
    const start = (await invoke('sync-ui:create-snapshot')) as { started: boolean };
    assert.equal(start.started, true);
    const args = launcherCalls[0]!.args;
    assert.equal(args[0], 'sync');
    assert.equal(args[1], '--snapshot');
    assert.ok(args[2]!.startsWith(path.join(root, 'snapshots')));
    assert.match(args[2]!, /immersion-.*\.sqlite$/);
    assert.equal(args[3], '--json');
    launcherCalls[0]!.finish({ ok: true, error: null });
  }));

test('delete-snapshot refuses paths outside the snapshots dir', () =>
  withTempDir(async (root) => {
    const { invoke } = makeTestRig(root);
    fs.mkdirSync(path.join(root, 'snapshots'), { recursive: true });
    const inside = path.join(root, 'snapshots', 'immersion-1.sqlite');
    fs.writeFileSync(inside, 'x');
    const outside = path.join(root, 'immersion.sqlite');
    fs.writeFileSync(outside, 'x');

    await assert.rejects(async () => invoke('sync-ui:delete-snapshot', outside));
    const remaining = (await invoke('sync-ui:delete-snapshot', inside)) as unknown[];
    assert.equal(remaining.length, 0);
    assert.equal(fs.existsSync(inside), false);
    assert.equal(fs.existsSync(outside), true);
  }));

test('check-host resolves with the check-result event', () =>
  withTempDir(async (root) => {
    const { invoke, launcherCalls } = makeTestRig(root);
    const pending = invoke('sync-ui:check-host', 'media-box') as Promise<{
      ok: boolean;
      sshOk: boolean;
    }>;
    await new Promise((resolve) => setImmediate(resolve));
    assert.deepEqual(launcherCalls[0]!.args, ['sync', 'media-box', '--check', '--json']);
    assert.equal(launcherCalls[0]!.timeoutMs, 30_000);
    launcherCalls[0]!.onEvent({
      type: 'check-result',
      host: 'media-box',
      sshOk: true,
      remoteCommand: 'subminer',
      remoteVersion: '0.18.0',
      ok: true,
      error: null,
    });
    launcherCalls[0]!.finish({ ok: true, error: null });
    const result = await pending;
    assert.equal(result.ok, true);
    assert.equal(result.sshOk, true);
  }));

test('check-host participates in active-run coordination and can be cancelled', () =>
  withTempDir(async (root) => {
    const { invoke, launcherCalls } = makeTestRig(root);
    const pending = invoke('sync-ui:check-host', 'media-box') as Promise<{ ok: boolean }>;
    await new Promise((resolve) => setImmediate(resolve));

    const blocked = (await invoke('sync-ui:run-sync', { host: 'other' })) as {
      started: boolean;
      reason: string | null;
    };
    assert.equal(blocked.started, false);
    assert.match(blocked.reason ?? '', /already running/i);

    assert.equal(await invoke('sync-ui:cancel-run'), true);
    assert.equal(launcherCalls[0]!.cancelled, true);
    assert.equal((await pending).ok, false);
  }));

test('shutdown cancels and awaits the active launcher run', () =>
  withTempDir(async (root) => {
    const { runtime, invoke, launcherCalls } = makeTestRig(root);
    await invoke('sync-ui:create-snapshot');
    await runtime.shutdown();
    assert.equal(launcherCalls[0]!.cancelled, true);
    assert.equal(runtime.isRunning(), false);
  }));

test('post-run side-effect failures are logged without rejecting cleanup', () =>
  withTempDir(async (root) => {
    const logs: string[] = [];
    const { invoke, launcherCalls } = makeTestRig(root, {
      getWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          send: () => {
            throw new Error('window gone');
          },
        },
      }),
      log: (message) => logs.push(message),
    });
    await invoke('sync-ui:create-snapshot');
    launcherCalls[0]!.finish({ ok: true, error: null });
    await new Promise((resolve) => setImmediate(resolve));
    assert.ok(logs.some((message) => message.includes('window gone')));
  }));

test('cancel-run cancels the active run', () =>
  withTempDir(async (root) => {
    const { invoke, launcherCalls } = makeTestRig(root);
    await invoke('sync-ui:run-sync', { host: 'media-box' });
    const cancelled = (await invoke('sync-ui:cancel-run')) as boolean;
    assert.equal(cancelled, true);
    assert.equal(launcherCalls[0]!.cancelled, true);
  }));

test('get-snapshot lists snapshot files newest first', () =>
  withTempDir(async (root) => {
    const { invoke } = makeTestRig(root);
    const dir = path.join(root, 'snapshots');
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, 'immersion-old.sqlite'), 'old');
    fs.writeFileSync(path.join(dir, 'ignore.txt'), 'x');
    fs.writeFileSync(path.join(dir, 'immersion-new.sqlite'), 'new');
    fs.utimesSync(path.join(dir, 'immersion-old.sqlite'), new Date(1000), new Date(1000));
    fs.utimesSync(path.join(dir, 'immersion-new.sqlite'), new Date(2000), new Date(2000));

    const snapshot = (await invoke('sync-ui:get-snapshot')) as {
      snapshots: Array<{ name: string }>;
      dbPath: string;
      run: { running: boolean };
    };
    assert.deepEqual(
      snapshot.snapshots.map((file) => file.name),
      ['immersion-new.sqlite', 'immersion-old.sqlite'],
    );
    assert.equal(snapshot.run.running, false);
  }));
