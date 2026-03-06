import test from 'node:test';
import assert from 'node:assert/strict';
import { createCriticalConfigErrorHandler, createReloadConfigHandler } from './startup-config';

test('createReloadConfigHandler runs success flow with warnings', async () => {
  const calls: string[] = [];
  const refreshCalls: { force: boolean }[] = [];

  const reloadConfig = createReloadConfigHandler({
    reloadConfigStrict: () => ({
      ok: true,
      path: '/tmp/config.jsonc',
      warnings: [
        {
          path: 'ankiConnect.pollingRate',
          message: 'must be >= 50',
          value: 10,
          fallback: 250,
        },
      ],
    }),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarning: (message) => calls.push(`warn:${message}`),
    showDesktopNotification: (title, options) => calls.push(`notify:${title}:${options.body}`),
    startConfigHotReload: () => calls.push('hotReload:start'),
    refreshAnilistClientSecretState: async (options) => {
      refreshCalls.push(options);
    },
    failHandlers: {
      logError: (details) => calls.push(`error:${details}`),
      showErrorBox: (title, details) => calls.push(`dialog:${title}:${details}`),
      quit: () => calls.push('quit'),
    },
  });

  reloadConfig();
  await Promise.resolve();

  assert.ok(calls.some((entry) => entry.startsWith('info:Using config file: /tmp/config.jsonc')));
  assert.ok(calls.some((entry) => entry.startsWith('warn:[config] Validation found 1 issue(s)')));
  assert.ok(
    calls.some((entry) => entry.includes('notify:SubMiner:1 config validation issue(s) detected.')),
  );
  assert.ok(calls.some((entry) => entry.includes('1. ankiConnect.pollingRate: must be >= 50')));
  const showedWarningDialog = calls.some((entry) =>
    entry.includes(
      'dialog:SubMiner config validation warning:SubMiner detected config validation issues.',
    ),
  );
  assert.equal(showedWarningDialog, process.platform === 'darwin');
  assert.ok(calls.some((entry) => entry.includes('actual=10 fallback=250')));
  assert.ok(calls.includes('hotReload:start'));
  assert.deepEqual(refreshCalls, [{ force: true }]);
});

test('createReloadConfigHandler fails startup for parse errors', () => {
  const calls: string[] = [];
  const previousExitCode = process.exitCode;
  process.exitCode = 0;

  const reloadConfig = createReloadConfigHandler({
    reloadConfigStrict: () => ({
      ok: false,
      path: '/tmp/config.jsonc',
      error: 'unexpected token',
    }),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarning: (message) => calls.push(`warn:${message}`),
    showDesktopNotification: (title, options) => calls.push(`notify:${title}:${options.body}`),
    startConfigHotReload: () => calls.push('hotReload:start'),
    refreshAnilistClientSecretState: async () => {
      calls.push('refresh');
    },
    failHandlers: {
      logError: (details) => calls.push(`error:${details}`),
      showErrorBox: (title, details) => calls.push(`dialog:${title}:${details}`),
      quit: () => calls.push('quit'),
    },
  });

  assert.throws(() => reloadConfig(), /Failed to parse config file at:/);
  assert.equal(process.exitCode, 1);
  assert.ok(calls.some((entry) => entry.startsWith('error:Failed to parse config file at:')));
  assert.ok(calls.some((entry) => entry.includes('/tmp/config.jsonc')));
  assert.ok(calls.some((entry) => entry.includes('Error: unexpected token')));
  assert.ok(calls.some((entry) => entry.includes('Fix the config file and restart SubMiner.')));
  assert.ok(
    calls.some((entry) =>
      entry.startsWith('dialog:SubMiner config parse error:Failed to parse config file at:'),
    ),
  );
  assert.ok(calls.includes('quit'));
  assert.equal(calls.includes('hotReload:start'), false);

  process.exitCode = previousExitCode;
});

test('createCriticalConfigErrorHandler formats and fails', () => {
  const calls: string[] = [];
  const previousExitCode = process.exitCode;
  process.exitCode = 0;

  const handleCriticalErrors = createCriticalConfigErrorHandler({
    getConfigPath: () => '/tmp/config.jsonc',
    failHandlers: {
      logError: (details) => calls.push(`error:${details}`),
      showErrorBox: (title, details) => calls.push(`dialog:${title}:${details}`),
      quit: () => calls.push('quit'),
    },
  });

  assert.throws(
    () => handleCriticalErrors(['foo invalid', 'bar invalid']),
    /Critical config validation failed/,
  );

  assert.equal(process.exitCode, 1);
  assert.ok(calls.some((entry) => entry.includes('/tmp/config.jsonc')));
  assert.ok(calls.some((entry) => entry.includes('1. foo invalid')));
  assert.ok(calls.some((entry) => entry.includes('2. bar invalid')));
  assert.ok(calls.includes('quit'));

  process.exitCode = previousExitCode;
});
