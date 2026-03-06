import test from 'node:test';
import assert from 'node:assert/strict';

async function loadRegistryOrSkip(t: test.TestContext) {
  try {
    return await import('./registry');
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes('node:sqlite')) {
      t.skip('registry import requires node:sqlite support in this runtime');
      return null;
    }
    throw error;
  }
}

test('createMainRuntimeRegistry exposes expected runtime domains', async (t) => {
  const loaded = await loadRegistryOrSkip(t);
  if (!loaded) return;
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

test('registry domains expose representative factories', async (t) => {
  const loaded = await loadRegistryOrSkip(t);
  if (!loaded) return;
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
