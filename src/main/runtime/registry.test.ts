import test from 'node:test';
import assert from 'node:assert/strict';

async function loadRegistry() {
  return import('./registry');
}

test('createMainRuntimeRegistry exposes expected runtime domains', async () => {
  const loaded = await loadRegistry();
  const { createMainRuntimeRegistry } = loaded;
  const registry = createMainRuntimeRegistry();

  assert.ok(registry.anilist);
  assert.ok(registry.jellyfin);
  assert.ok(registry.overlay);
  assert.ok(registry.startup);
  assert.ok(registry.mpv);
  assert.ok(registry.shortcuts);
  assert.ok(registry.ipc);
  assert.ok(registry.mining);
});

test('registry domains expose representative factories', async () => {
  const loaded = await loadRegistry();
  const { createMainRuntimeRegistry } = loaded;
  const registry = createMainRuntimeRegistry();

  assert.equal(typeof registry.anilist.createNotifyAnilistSetupHandler, 'function');
  assert.equal(typeof registry.jellyfin.createRunJellyfinCommandHandler, 'function');
  assert.equal(typeof registry.overlay.createOverlayVisibilityRuntime, 'function');
  assert.equal(typeof registry.startup.createStartupRuntimeHandlers, 'function');
  assert.equal(typeof registry.mpv.createMpvClientRuntimeServiceFactory, 'function');
  assert.equal(typeof registry.shortcuts.createGlobalShortcutsRuntimeHandlers, 'function');
  assert.equal(typeof registry.ipc.createIpcRuntimeHandlers, 'function');
  assert.equal(typeof registry.mining.createMineSentenceCardHandler, 'function');
});
