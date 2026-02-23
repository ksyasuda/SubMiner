import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildStartupBootstrapMainDepsHandler } from './startup-bootstrap-main-deps';

test('startup bootstrap main deps builder maps deps and handles generate-config callbacks', () => {
  const calls: string[] = [];
  let exitCode = 0;
  const deps = createBuildStartupBootstrapMainDepsHandler({
    argv: ['node', 'main.js'],
    parseArgs: () => ({}) as never,
    setLogLevel: (level) => calls.push(`log:${level}`),
    forceX11Backend: () => calls.push('force-x11'),
    enforceUnsupportedWaylandMode: () => calls.push('guard-wayland'),
    shouldStartApp: () => true,
    getDefaultSocketPath: () => '/tmp/mpv.sock',
    defaultTexthookerPort: 5174,
    configDir: '/tmp/config',
    defaultConfig: {} as never,
    generateConfigTemplate: () => 'template',
    generateDefaultConfigFile: async () => 0,
    setExitCode: (code) => {
      exitCode = code;
      calls.push(`exit:${code}`);
    },
    quitApp: () => calls.push('quit'),
    logGenerateConfigError: (message) => calls.push(`error:${message}`),
    startAppLifecycle: () => calls.push('start-lifecycle'),
  })();

  assert.deepEqual(deps.argv, ['node', 'main.js']);
  assert.equal(deps.getDefaultSocketPath(), '/tmp/mpv.sock');
  deps.setLogLevel('debug', 'config');
  deps.forceX11Backend({} as never);
  deps.enforceUnsupportedWaylandMode({} as never);
  deps.startAppLifecycle({} as never);
  deps.onConfigGenerated(7);
  assert.equal(exitCode, 7);
  deps.onGenerateConfigError(new Error('boom'));
  assert.equal(exitCode, 1);

  assert.deepEqual(calls, [
    'log:debug',
    'force-x11',
    'guard-wayland',
    'start-lifecycle',
    'exit:7',
    'quit',
    'error:Failed to generate config: boom',
    'exit:1',
    'quit',
  ]);
});
