import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildStartupBootstrapMainDepsHandler } from './startup-bootstrap-main-deps';

function buildDeps() {
  const state = { exitCodes: [] as number[], quitCount: 0, errors: [] as string[] };
  const deps = createBuildStartupBootstrapMainDepsHandler({
    argv: ['node', 'main.js'],
    parseArgs: () => ({}) as never,
    setLogLevel: () => {},
    forceX11Backend: () => {},
    enforceUnsupportedWaylandMode: () => {},
    shouldStartApp: () => true,
    getDefaultSocketPath: () => '/tmp/mpv.sock',
    defaultTexthookerPort: 5174,
    configDir: '/tmp/config',
    defaultConfig: {} as never,
    generateConfigTemplate: () => 'template',
    generateDefaultConfigFile: async () => 0,
    setExitCode: (code) => state.exitCodes.push(code),
    quitApp: () => {
      state.quitCount += 1;
    },
    logGenerateConfigError: (message) => state.errors.push(message),
    startAppLifecycle: () => {},
  })();
  return { deps, state };
}

test('onConfigGenerated exits with the generator exit code', () => {
  const { deps, state } = buildDeps();
  deps.onConfigGenerated(7);
  assert.deepEqual(state.exitCodes, [7]);
  assert.equal(state.quitCount, 1);
  assert.deepEqual(state.errors, []);
});

test('onGenerateConfigError logs the failure and exits with code 1', () => {
  const { deps, state } = buildDeps();
  deps.onGenerateConfigError(new Error('boom'));
  assert.deepEqual(state.errors, ['Failed to generate config: boom']);
  assert.deepEqual(state.exitCodes, [1]);
  assert.equal(state.quitCount, 1);
});
