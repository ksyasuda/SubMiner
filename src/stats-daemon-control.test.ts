import assert from 'node:assert/strict';
import test from 'node:test';
import { createRunStatsDaemonControlHandler } from './stats-daemon-control';

test('stats daemon control reuses live daemon and writes launcher response', async () => {
  const calls: string[] = [];
  const responses: Array<{ path: string; payload: { ok: boolean; url?: string; error?: string } }> =
    [];
  const handler = createRunStatsDaemonControlHandler({
    statePath: '/tmp/stats-daemon.json',
    readState: () => ({ pid: 4242, port: 5175, startedAtMs: 1 }),
    removeState: () => {
      calls.push('removeState');
    },
    isProcessAlive: (pid) => {
      calls.push(`isProcessAlive:${pid}`);
      return true;
    },
    resolveUrl: (state) => `http://127.0.0.1:${state.port}`,
    spawnDaemon: async () => {
      calls.push('spawnDaemon');
      return 1;
    },
    waitForDaemonResponse: async () => {
      calls.push('waitForDaemonResponse');
      return { ok: true, url: 'http://127.0.0.1:5175' };
    },
    openExternal: async (url) => {
      calls.push(`openExternal:${url}`);
    },
    writeResponse: (responsePath, payload) => {
      responses.push({ path: responsePath, payload });
    },
    killProcess: () => {
      calls.push('killProcess');
    },
    sleep: async () => {},
  });

  const exitCode = await handler({
    action: 'start',
    responsePath: '/tmp/response.json',
    openBrowser: true,
    daemonScriptPath: '/tmp/stats-daemon-runner.js',
    userDataPath: '/tmp/SubMiner',
  });

  assert.equal(exitCode, 0);
  assert.deepEqual(calls, ['isProcessAlive:4242', 'openExternal:http://127.0.0.1:5175']);
  assert.deepEqual(responses, [
    {
      path: '/tmp/response.json',
      payload: { ok: true, url: 'http://127.0.0.1:5175' },
    },
  ]);
});

test('stats daemon control clears stale state, starts daemon, and waits for response', async () => {
  const calls: string[] = [];
  const handler = createRunStatsDaemonControlHandler({
    statePath: '/tmp/stats-daemon.json',
    readState: () => ({ pid: 4242, port: 5175, startedAtMs: 1 }),
    removeState: () => {
      calls.push('removeState');
    },
    isProcessAlive: (pid) => {
      calls.push(`isProcessAlive:${pid}`);
      return false;
    },
    resolveUrl: (state) => `http://127.0.0.1:${state.port}`,
    spawnDaemon: async (options) => {
      calls.push(
        `spawnDaemon:${options.scriptPath}:${options.responsePath}:${options.userDataPath}`,
      );
      return 999;
    },
    waitForDaemonResponse: async (responsePath) => {
      calls.push(`waitForDaemonResponse:${responsePath}`);
      return { ok: true, url: 'http://127.0.0.1:5175' };
    },
    openExternal: async (url) => {
      calls.push(`openExternal:${url}`);
    },
    writeResponse: () => {
      calls.push('writeResponse');
    },
    killProcess: () => {
      calls.push('killProcess');
    },
    sleep: async () => {},
  });

  const exitCode = await handler({
    action: 'start',
    responsePath: '/tmp/response.json',
    openBrowser: false,
    daemonScriptPath: '/tmp/stats-daemon-runner.js',
    userDataPath: '/tmp/SubMiner',
  });

  assert.equal(exitCode, 0);
  assert.deepEqual(calls, [
    'isProcessAlive:4242',
    'removeState',
    'spawnDaemon:/tmp/stats-daemon-runner.js:/tmp/response.json:/tmp/SubMiner',
    'waitForDaemonResponse:/tmp/response.json',
  ]);
});

test('stats daemon control stops live daemon and treats stale state as success', async () => {
  const responses: Array<{ path: string; payload: { ok: boolean; url?: string; error?: string } }> =
    [];
  const calls: string[] = [];
  let aliveChecks = 0;
  const handler = createRunStatsDaemonControlHandler({
    statePath: '/tmp/stats-daemon.json',
    readState: () => ({ pid: 4242, port: 5175, startedAtMs: 1 }),
    removeState: () => {
      calls.push('removeState');
    },
    isProcessAlive: (pid) => {
      aliveChecks += 1;
      calls.push(`isProcessAlive:${pid}:${aliveChecks}`);
      return aliveChecks === 1;
    },
    verifyProcessIdentity: (pid, startedAtMs) => {
      calls.push(`verifyProcessIdentity:${pid}:${startedAtMs}`);
      return true;
    },
    resolveUrl: (state) => `http://127.0.0.1:${state.port}`,
    spawnDaemon: async () => 1,
    waitForDaemonResponse: async () => ({ ok: true, url: 'http://127.0.0.1:5175' }),
    openExternal: async () => {},
    writeResponse: (responsePath, payload) => {
      responses.push({ path: responsePath, payload });
    },
    killProcess: (pid, signal) => {
      calls.push(`killProcess:${pid}:${signal}`);
    },
    sleep: async () => {},
  });

  const exitCode = await handler({
    action: 'stop',
    responsePath: '/tmp/response.json',
    openBrowser: false,
    daemonScriptPath: '/tmp/stats-daemon-runner.js',
    userDataPath: '/tmp/SubMiner',
  });

  assert.equal(exitCode, 0);
  assert.deepEqual(calls, [
    'isProcessAlive:4242:1',
    'verifyProcessIdentity:4242:1',
    'killProcess:4242:SIGTERM',
    'isProcessAlive:4242:2',
    'removeState',
  ]);
  assert.deepEqual(responses, [
    {
      path: '/tmp/response.json',
      payload: { ok: true },
    },
  ]);
});

test('stats daemon control clears stale state when daemon identity mismatches', async () => {
  const calls: string[] = [];
  const responses: Array<{ path: string; payload: { ok: boolean; url?: string; error?: string } }> =
    [];
  const handler = createRunStatsDaemonControlHandler({
    statePath: '/tmp/stats-daemon.json',
    readState: () => ({ pid: 4242, port: 5175, startedAtMs: 1 }),
    removeState: () => {
      calls.push('removeState');
    },
    isProcessAlive: (pid) => {
      calls.push(`isProcessAlive:${pid}`);
      return true;
    },
    verifyProcessIdentity: (pid, startedAtMs) => {
      calls.push(`verifyProcessIdentity:${pid}:${startedAtMs}`);
      return false;
    },
    resolveUrl: (state) => `http://127.0.0.1:${state.port}`,
    spawnDaemon: async () => 1,
    waitForDaemonResponse: async () => ({ ok: true, url: 'http://127.0.0.1:5175' }),
    openExternal: async () => {},
    writeResponse: (responsePath, payload) => {
      responses.push({ path: responsePath, payload });
    },
    killProcess: () => {
      calls.push('killProcess');
    },
    sleep: async () => {},
  });

  const exitCode = await handler({
    action: 'stop',
    responsePath: '/tmp/response.json',
    openBrowser: false,
    daemonScriptPath: '/tmp/stats-daemon-runner.js',
    userDataPath: '/tmp/SubMiner',
  });

  assert.equal(exitCode, 0);
  assert.deepEqual(calls, ['isProcessAlive:4242', 'verifyProcessIdentity:4242:1', 'removeState']);
  assert.deepEqual(responses, [{ path: '/tmp/response.json', payload: { ok: true } }]);
});
