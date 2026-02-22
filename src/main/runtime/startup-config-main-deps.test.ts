import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildCriticalConfigErrorMainDepsHandler,
  createBuildReloadConfigMainDepsHandler,
} from './startup-config-main-deps';

test('reload config main deps builder maps callbacks and fail handlers', async () => {
  const calls: string[] = [];
  const deps = createBuildReloadConfigMainDepsHandler({
    reloadConfigStrict: () => ({ ok: true, path: '/tmp/config.jsonc', warnings: [] }),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarning: (message) => calls.push(`warn:${message}`),
    showDesktopNotification: (title, options) => calls.push(`notify:${title}:${options.body}`),
    startConfigHotReload: () => calls.push('start-hot-reload'),
    refreshAnilistClientSecretState: async (options) => {
      calls.push(`refresh:${options.force}`);
      return true;
    },
    failHandlers: {
      logError: (details) => calls.push(`error:${details}`),
      showErrorBox: (title, details) => calls.push(`error-box:${title}:${details}`),
      quit: () => calls.push('quit'),
    },
  })();

  assert.deepEqual(deps.reloadConfigStrict(), {
    ok: true,
    path: '/tmp/config.jsonc',
    warnings: [],
  });
  deps.logInfo('x');
  deps.logWarning('y');
  deps.showDesktopNotification('SubMiner', { body: 'warn' });
  deps.startConfigHotReload();
  await deps.refreshAnilistClientSecretState({ force: true });
  deps.failHandlers.logError('bad');
  deps.failHandlers.showErrorBox('Oops', 'Details');
  deps.failHandlers.quit();
  assert.deepEqual(calls, [
    'info:x',
    'warn:y',
    'notify:SubMiner:warn',
    'start-hot-reload',
    'refresh:true',
    'error:bad',
    'error-box:Oops:Details',
    'quit',
  ]);
});

test('critical config main deps builder maps config path and fail handlers', () => {
  const calls: string[] = [];
  const deps = createBuildCriticalConfigErrorMainDepsHandler({
    getConfigPath: () => '/tmp/config.jsonc',
    failHandlers: {
      logError: (details) => calls.push(`error:${details}`),
      showErrorBox: (title, details) => calls.push(`error-box:${title}:${details}`),
      quit: () => calls.push('quit'),
    },
  })();

  assert.equal(deps.getConfigPath(), '/tmp/config.jsonc');
  deps.failHandlers.logError('bad');
  deps.failHandlers.showErrorBox('Oops', 'Details');
  deps.failHandlers.quit();
  assert.deepEqual(calls, ['error:bad', 'error-box:Oops:Details', 'quit']);
});
