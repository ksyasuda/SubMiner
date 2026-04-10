import assert from 'node:assert/strict';
import test from 'node:test';
import type { ReloadConfigStrictResult } from '../../config';
import type { ResolvedConfig } from '../../types';
import {
  createBuildConfigHotReloadMessageMainDepsHandler,
  createBuildConfigHotReloadAppliedMainDepsHandler,
  createBuildConfigHotReloadRuntimeMainDepsHandler,
  createBuildWatchConfigPathMainDepsHandler,
  createWatchConfigPathHandler,
} from './config-hot-reload-main-deps';

test('watch config path handler watches file directly when config exists', () => {
  const calls: string[] = [];
  const watchConfigPath = createWatchConfigPathHandler({
    fileExists: () => true,
    dirname: (path) => path.split('/').slice(0, -1).join('/'),
    watchPath: (targetPath, nextListener) => {
      calls.push(`watch:${targetPath}`);
      nextListener('change', 'ignored');
      return { close: () => calls.push('close') };
    },
  });

  const watcher = watchConfigPath('/tmp/config.jsonc', () => calls.push('change'));
  watcher.close();
  assert.deepEqual(calls, ['watch:/tmp/config.jsonc', 'change', 'close']);
});

test('watch config path handler filters directory events to config files only', () => {
  const calls: string[] = [];
  const watchConfigPath = createWatchConfigPathHandler({
    fileExists: () => false,
    dirname: (path) => path.split('/').slice(0, -1).join('/'),
    watchPath: (_targetPath, nextListener) => {
      nextListener('change', 'foo.txt');
      nextListener('change', 'config.json');
      nextListener('change', 'config.jsonc');
      nextListener('change', null);
      return { close: () => {} };
    },
  });

  watchConfigPath('/tmp/config.jsonc', () => calls.push('change'));
  assert.deepEqual(calls, ['change', 'change', 'change']);
});

test('watch config path main deps builder maps filesystem callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildWatchConfigPathMainDepsHandler({
    fileExists: () => true,
    dirname: (targetPath) => {
      calls.push(`dirname:${targetPath}`);
      return '/tmp';
    },
    watchPath: (targetPath, listener) => {
      calls.push(`watch:${targetPath}`);
      listener('change', 'config.jsonc');
      return { close: () => calls.push('close') };
    },
  })();

  assert.equal(deps.fileExists('/tmp/config.jsonc'), true);
  assert.equal(deps.dirname('/tmp/config.jsonc'), '/tmp');
  const watcher = deps.watchPath('/tmp/config.jsonc', () => calls.push('listener'));
  watcher.close();
  assert.deepEqual(calls, [
    'dirname:/tmp/config.jsonc',
    'watch:/tmp/config.jsonc',
    'listener',
    'close',
  ]);
});

test('config hot reload message main deps builder maps notifications', () => {
  const calls: string[] = [];
  const deps = createBuildConfigHotReloadMessageMainDepsHandler({
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    showDesktopNotification: (title) => calls.push(`notify:${title}`),
  })();

  deps.showMpvOsd('updated');
  deps.showDesktopNotification('SubMiner', { body: 'updated' });
  assert.deepEqual(calls, ['osd:updated', 'notify:SubMiner']);
});

test('config hot reload applied main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const buildDeps = createBuildConfigHotReloadAppliedMainDepsHandler({
    setKeybindings: () => calls.push('keybindings'),
    setSessionBindings: () => calls.push('session-bindings'),
    refreshGlobalAndOverlayShortcuts: () => calls.push('refresh-shortcuts'),
    setSecondarySubMode: () => calls.push('set-secondary'),
    broadcastToOverlayWindows: (channel) => calls.push(`broadcast:${channel}`),
    applyAnkiRuntimeConfigPatch: () => calls.push('apply-anki'),
  });

  const deps = buildDeps();
  deps.setKeybindings([]);
  deps.setSessionBindings([]);
  deps.refreshGlobalAndOverlayShortcuts();
  deps.setSecondarySubMode('hover');
  deps.broadcastToOverlayWindows('config:hot-reload', {});
  deps.applyAnkiRuntimeConfigPatch({ ai: true });
  assert.deepEqual(calls, [
    'keybindings',
    'session-bindings',
    'refresh-shortcuts',
    'set-secondary',
    'broadcast:config:hot-reload',
    'apply-anki',
  ]);
});

test('config hot reload runtime main deps builder maps runtime callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildConfigHotReloadRuntimeMainDepsHandler({
    getCurrentConfig: () => ({ id: 1 }) as never as ResolvedConfig,
    reloadConfigStrict: () =>
      ({
        ok: true,
        config: { id: 1 } as never as ResolvedConfig,
        warnings: [],
        path: '/tmp/config.jsonc',
      }) as ReloadConfigStrictResult,
    watchConfigPath: (_configPath, _onChange) => ({ close: () => calls.push('close') }),
    setTimeout: (callback) => {
      callback();
      return 1 as never;
    },
    clearTimeout: () => calls.push('clear-timeout'),
    debounceMs: 250,
    onHotReloadApplied: () => calls.push('hot-reload'),
    onRestartRequired: () => calls.push('restart-required'),
    onInvalidConfig: () => calls.push('invalid-config'),
    onValidationWarnings: () => calls.push('validation-warnings'),
  })();

  assert.deepEqual(deps.getCurrentConfig(), { id: 1 });
  assert.deepEqual(deps.reloadConfigStrict(), {
    ok: true,
    config: { id: 1 },
    warnings: [],
    path: '/tmp/config.jsonc',
  });
  assert.equal(deps.debounceMs, 250);
  deps.onHotReloadApplied({} as never, {} as never);
  deps.onRestartRequired([]);
  deps.onInvalidConfig('bad');
  deps.onValidationWarnings('/tmp/config.jsonc', []);
  assert.deepEqual(calls, [
    'hot-reload',
    'restart-required',
    'invalid-config',
    'validation-warnings',
  ]);
});
