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

test('persistSessionBindings keeps saved bindings when mpv reload notification fails', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-session-bindings-runtime-'));
  const calls: string[] = [];
  const runtime = createSessionBindingsRuntime({
    configDir: root,
    getKeybindings: () => [],
    getConfiguredShortcuts: () => ({ multiCopyTimeoutMs: 1500 }) as never,
    getResolvedConfig: () =>
      ({
        stats: { toggleKey: 's', markWatchedKey: 'w' },
      }) as ResolvedConfig,
    getMpvClient: () =>
      ({
        connected: true,
        send: () => {
          throw new Error('mpv unavailable');
        },
      }) as never,
    setSessionBindings: () => calls.push('setSessionBindings'),
    setSessionBindingsInitialized: () => calls.push('setSessionBindingsInitialized'),
    logWarn: (message) => calls.push(`warn:${message}`),
  });

  try {
    assert.doesNotThrow(() => runtime.persistSessionBindings([] as CompiledSessionBinding[]));
    assert.deepEqual(calls, [
      'setSessionBindings',
      'setSessionBindingsInitialized',
      'warn:[session-bindings] Failed to notify mpv to reload session bindings',
    ]);
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('native prefix conflicts publish the same effective bindings to the overlay and plugin and recover', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-session-conflict-'));
  const sequence: CompiledSessionBinding = {
    sourcePath: 'shortcuts.openSubtitleSelection',
    originalKey: 'g-s',
    key: { code: 'KeyG-KeyS', modifiers: [] },
    actionType: 'session-action',
    actionId: 'openSubtitleSelection',
  };
  let nativeKeys: unknown = [];
  let failDiscovery = false;
  let published: CompiledSessionBinding[] = [];
  const events: CompiledSessionBinding[][] = [];
  const warnings: string[] = [];
  const client = {
    connected: true,
    send: () => {},
    requestProperty: async () => {
      if (failDiscovery) throw new Error('temporarily unavailable');
      return nativeKeys;
    },
  };
  const runtime = createSessionBindingsRuntime({
    configDir: root,
    getKeybindings: () => [],
    getConfiguredShortcuts: () => ({ multiCopyTimeoutMs: 1500 }) as never,
    getResolvedConfig: () => ({ stats: { toggleKey: 's', markWatchedKey: 'w' } }) as ResolvedConfig,
    getMpvClient: () => client,
    setSessionBindings: (bindings) => {
      published = bindings;
    },
    setSessionBindingsInitialized: () => {},
    logWarn: () => {},
    onBindingsChanged: (bindings) => events.push(bindings),
    onWarning: (warning) => warnings.push(warning.message),
  });
  const readArtifact = () =>
    JSON.parse(fs.readFileSync(path.join(root, 'session-bindings.json'), 'utf8'));
  try {
    runtime.persistSessionBindings([sequence]);
    nativeKeys = [{ key: 'g', cmd: 'show-text single', priority: 1 }];
    await runtime.refreshMpvSessionBindings();
    assert.deepEqual(published, []);
    assert.deepEqual(events.at(-1), readArtifact().bindings);
    assert.equal(warnings.length, 1);
    assert.match(warnings[0]!, /mpv input binding "g"/);
    await runtime.refreshMpvSessionBindings();
    assert.equal(events.length, 2, 'unchanged discovery must not create a reload loop');
    assert.equal(warnings.length, 1);
    failDiscovery = true;
    await runtime.refreshMpvSessionBindings();
    assert.deepEqual(published, [], 'failed discovery retains the known conflict');
    failDiscovery = false;
    nativeKeys = [{ key: 'Shift+g', cmd: 'show-text shifted', priority: 1 }];
    await runtime.refreshMpvSessionBindings();
    assert.deepEqual(published, [sequence]);
    assert.equal(readArtifact().bindings[0].key.code, 'KeyG-KeyS');
    assert.deepEqual(readArtifact().warnings, []);
    client.connected = false;
    await runtime.refreshMpvSessionBindings();
    assert.deepEqual(published, [sequence]);
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});
