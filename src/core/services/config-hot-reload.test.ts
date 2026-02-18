import test from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_CONFIG, deepCloneConfig } from '../../config';
import {
  classifyConfigHotReloadDiff,
  createConfigHotReloadRuntime,
  type ConfigHotReloadRuntimeDeps,
} from './config-hot-reload';

test('classifyConfigHotReloadDiff separates hot and restart-required fields', () => {
  const prev = deepCloneConfig(DEFAULT_CONFIG);
  const next = deepCloneConfig(DEFAULT_CONFIG);
  next.subtitleStyle.fontSize = prev.subtitleStyle.fontSize + 2;
  next.websocket.port = prev.websocket.port + 1;

  const diff = classifyConfigHotReloadDiff(prev, next);
  assert.deepEqual(diff.hotReloadFields, ['subtitleStyle']);
  assert.deepEqual(diff.restartRequiredFields, ['websocket']);
});

test('config hot reload runtime debounces rapid watch events', () => {
  let watchedChangeCallback: (() => void) | null = null;
  const pendingTimers = new Map<number, () => void>();
  let nextTimerId = 1;
  let reloadCalls = 0;

  const deps: ConfigHotReloadRuntimeDeps = {
    getCurrentConfig: () => deepCloneConfig(DEFAULT_CONFIG),
    reloadConfigStrict: () => {
      reloadCalls += 1;
      return {
        ok: true,
        config: deepCloneConfig(DEFAULT_CONFIG),
        warnings: [],
        path: '/tmp/config.jsonc',
      };
    },
    watchConfigPath: (_path, onChange) => {
      watchedChangeCallback = onChange;
      return { close: () => {} };
    },
    setTimeout: (callback) => {
      const id = nextTimerId;
      nextTimerId += 1;
      pendingTimers.set(id, callback);
      return id as unknown as NodeJS.Timeout;
    },
    clearTimeout: (timeout) => {
      pendingTimers.delete(timeout as unknown as number);
    },
    debounceMs: 25,
    onHotReloadApplied: () => {},
    onRestartRequired: () => {},
    onInvalidConfig: () => {},
  };

  const runtime = createConfigHotReloadRuntime(deps);
  runtime.start();
  assert.equal(reloadCalls, 1);
  if (!watchedChangeCallback) {
    throw new Error('Expected watch callback to be registered.');
  }
  const trigger = watchedChangeCallback as () => void;

  trigger();
  trigger();
  trigger();
  assert.equal(pendingTimers.size, 1);

  for (const callback of pendingTimers.values()) {
    callback();
  }
  assert.equal(reloadCalls, 2);
});

test('config hot reload runtime reports invalid config and skips apply', () => {
  const invalidMessages: string[] = [];
  let watchedChangeCallback: (() => void) | null = null;

  const runtime = createConfigHotReloadRuntime({
    getCurrentConfig: () => deepCloneConfig(DEFAULT_CONFIG),
    reloadConfigStrict: () => ({
      ok: false,
      error: 'Invalid JSON',
      path: '/tmp/config.jsonc',
    }),
    watchConfigPath: (_path, onChange) => {
      watchedChangeCallback = onChange;
      return { close: () => {} };
    },
    setTimeout: (callback) => {
      callback();
      return 1 as unknown as NodeJS.Timeout;
    },
    clearTimeout: () => {},
    debounceMs: 0,
    onHotReloadApplied: () => {
      throw new Error('Hot reload should not apply for invalid config.');
    },
    onRestartRequired: () => {
      throw new Error('Restart warning should not trigger for invalid config.');
    },
    onInvalidConfig: (message) => {
      invalidMessages.push(message);
    },
  });

  runtime.start();
  assert.equal(watchedChangeCallback, null);
  assert.equal(invalidMessages.length, 1);
});
