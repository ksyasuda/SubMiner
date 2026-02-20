import test from 'node:test';
import assert from 'node:assert/strict';
import { createBuildStartupBootstrapRuntimeFactoryDepsHandler } from './startup-bootstrap-deps-builder';

test('startup bootstrap deps builder returns mapped runtime factory deps', () => {
  const calls: string[] = [];
  const factory = createBuildStartupBootstrapRuntimeFactoryDepsHandler({
    argv: ['node', 'main.js'],
    parseArgs: () => ({}) as never,
    setLogLevel: (level) => calls.push(`log:${level}`),
    forceX11Backend: () => calls.push('force-x11'),
    enforceUnsupportedWaylandMode: () => calls.push('wayland-guard'),
    shouldStartApp: () => true,
    getDefaultSocketPath: () => '/tmp/mpv.sock',
    defaultTexthookerPort: 5174,
    configDir: '/tmp/config',
    defaultConfig: {} as never,
    generateConfigTemplate: () => 'template',
    generateDefaultConfigFile: async () => 0,
    onConfigGenerated: (exitCode) => calls.push(`generated:${exitCode}`),
    onGenerateConfigError: (error) => calls.push(`error:${error.message}`),
    startAppLifecycle: () => calls.push('start-lifecycle'),
  });

  const deps = factory();
  assert.deepEqual(deps.argv, ['node', 'main.js']);
  assert.equal(deps.getDefaultSocketPath(), '/tmp/mpv.sock');
  assert.equal(deps.defaultTexthookerPort, 5174);
  deps.setLogLevel('debug', 'config');
  deps.forceX11Backend({} as never);
  deps.enforceUnsupportedWaylandMode({} as never);
  deps.onConfigGenerated(0);
  deps.onGenerateConfigError(new Error('oops'));
  deps.startAppLifecycle({} as never);

  assert.deepEqual(calls, [
    'log:debug',
    'force-x11',
    'wayland-guard',
    'generated:0',
    'error:oops',
    'start-lifecycle',
  ]);
});
