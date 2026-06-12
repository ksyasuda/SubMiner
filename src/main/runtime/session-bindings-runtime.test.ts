import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import type { CompiledSessionBinding, ResolvedConfig } from '../../types';
import { createSessionBindingsRuntime } from './session-bindings-runtime';

test('persistSessionBindings logs and does not publish bindings when artifact write fails', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-session-bindings-runtime-'));
  const configDir = path.join(root, 'config-file');
  fs.writeFileSync(configDir, 'not a directory');
  const calls: string[] = [];
  const runtime = createSessionBindingsRuntime({
    configDir,
    getKeybindings: () => [],
    getConfiguredShortcuts: () => ({ multiCopyTimeoutMs: 1500 }) as never,
    getResolvedConfig: () =>
      ({
        stats: { toggleKey: 's', markWatchedKey: 'w' },
      }) as ResolvedConfig,
    getMpvClient: () => null,
    setSessionBindings: () => calls.push('setSessionBindings'),
    setSessionBindingsInitialized: () => calls.push('setSessionBindingsInitialized'),
    logWarn: (message) => calls.push(`warn:${message}`),
  });

  try {
    assert.throws(
      () => runtime.persistSessionBindings([] as CompiledSessionBinding[]),
      /ENOTDIR|EEXIST/,
    );
    assert.deepEqual(calls, ['warn:[session-bindings] Failed to write session bindings artifact']);
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});
