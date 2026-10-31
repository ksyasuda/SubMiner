import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import type { BrowserWindow } from 'electron';

import type { SessionActionDispatchRequest } from '../../../types/runtime';
import { createPluginHost, type PluginHostDeps } from './plugin-host';
import type {
  PluginConfigEntry,
  PluginContext,
  PluginLogger,
  PluginMpvEventName,
  PluginMpvEventMap,
  PluginWindowRequest,
} from './types';

function writePlugin(root: string, id: string, files: Record<string, string>): void {
  const dir = path.join(root, id);
  mkdirSync(dir, { recursive: true });
  for (const [name, content] of Object.entries(files)) {
    writeFileSync(path.join(dir, name), content);
  }
}

type FakeWindow = ReturnType<typeof createFakeWindow>;

/** The slice of BrowserWindow the host touches: destroy, isDestroyed, and the `closed` event. */
function createFakeWindow() {
  const events = new EventEmitter();
  const fake = {
    destroyed: false,
    isDestroyed: () => fake.destroyed,
    destroy: () => {
      fake.destroyed = true;
      events.emit('closed');
    },
    once: (event: string, listener: () => void) => events.once(event, listener),
    asWindow: () => fake as unknown as BrowserWindow,
  };
  return fake;
}

function createHarness(
  config: Record<string, PluginConfigEntry>,
  overrides: Partial<PluginHostDeps<{ marker: string }>> = {},
) {
  const root = mkdtempSync(path.join(os.tmpdir(), 'subminer-plugins-'));
  const errors: string[] = [];
  const notifications: string[] = [];
  const sessionActions: SessionActionDispatchRequest[] = [];
  const windowRequests: PluginWindowRequest[] = [];
  const windows = new Map<string, FakeWindow>();
  const logger: PluginLogger = {
    debug: () => {},
    info: () => {},
    warn: () => {},
    error: (message) => errors.push(message),
  };
  const deps: PluginHostDeps<{ marker: string }> = {
    pluginsDir: path.join(root, 'plugins'),
    pluginDataDir: path.join(root, 'data'),
    getPluginConfig: () => config,
    createLogger: () => logger,
    isMpvConnected: () => true,
    sendMpvCommand: () => {},
    requestMpvProperty: async () => null,
    requestMpvCommand: async () => null,
    ensureMpvRunning: async () => true,
    invokeAnkiConnect: async () => null,
    dispatchSessionAction: async (request) => {
      sessionActions.push(request);
    },
    getCurrentSubtitle: () => null,
    getSubtitleContext: () => ({ current: null, before: [], after: [] }),
    notify: (message) => notifications.push(message),
    openPanel: () => {
      throw new Error('not used');
    },
    openWindow: (request) => {
      windowRequests.push(request);
      const existing = windows.get(request.key);
      if (existing && !existing.destroyed) return { window: existing.asWindow(), created: false };
      const window = createFakeWindow();
      windows.set(request.key, window);
      return { window: window.asWindow(), created: true };
    },
    unstable: { marker: 'internals' },
    ...overrides,
  };
  return {
    root: deps.pluginsDir,
    dataRoot: deps.pluginDataDir,
    errors,
    notifications,
    sessionActions,
    windowRequests,
    windows,
    host: createPluginHost(deps),
    cleanup: () => rmSync(root, { recursive: true, force: true }),
  };
}

const enabled = (settings: Record<string, unknown> = {}): PluginConfigEntry => ({
  enabled: true,
  settings,
});

function deferred() {
  let resolve = () => {};
  const promise = new Promise<void>((done) => {
    resolve = done;
  });
  return { promise, resolve: () => resolve() };
}

function createFakeMpvClient() {
  const bus = new EventEmitter();
  return {
    on: <E extends PluginMpvEventName>(
      event: E,
      listener: (payload: PluginMpvEventMap[E]) => void,
    ) => bus.on(event, listener),
    off: <E extends PluginMpvEventName>(
      event: E,
      listener: (payload: PluginMpvEventMap[E]) => void,
    ) => bus.off(event, listener),
    emit: <E extends PluginMpvEventName>(event: E, payload: PluginMpvEventMap[E]) =>
      bus.emit(event, payload),
  };
}

test('loads enabled plugins and runs their registered actions', async () => {
  const h = createHarness({
    greeter: enabled({ greeting: 'hi' }),
    off: { enabled: false, settings: {} },
  });
  try {
    writePlugin(h.root, 'greeter', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => {
        ctx.actions.register('say', () => ctx.notify(ctx.settings.greeting + ' ' + ctx.unstable.marker));
      };`,
    });
    writePlugin(h.root, 'off', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `throw new Error('disabled plugins must not load');`,
    });

    await h.host.loadAll();

    assert.equal(await h.host.runAction('greeter/say'), true);
    assert.deepEqual(h.notifications, ['hi internals']);
    assert.equal(await h.host.runAction('greeter/missing'), false);
    assert.deepEqual(h.errors, []);
  } finally {
    h.cleanup();
  }
});

test('loads ES module plugins', async () => {
  const h = createHarness({ esm: enabled() });
  try {
    writePlugin(h.root, 'esm', {
      'plugin.json': JSON.stringify({ main: 'main.mjs' }),
      'main.mjs': `export function activate(ctx) { ctx.actions.register('go', () => ctx.notify('esm')); }`,
    });

    await h.host.loadAll();

    assert.equal(await h.host.runAction('esm/go'), true);
    assert.deepEqual(h.notifications, ['esm']);
  } finally {
    h.cleanup();
  }
});

test('a failed activation rolls back its registrations without affecting other plugins', async () => {
  const h = createHarness({ broken: enabled(), healthy: enabled() });
  try {
    writePlugin(h.root, 'broken', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => {
        ctx.actions.register('half', () => {});
        throw new Error('boom');
      };`,
    });
    writePlugin(h.root, 'healthy', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => { ctx.actions.register('ok', () => {}); };`,
    });

    await h.host.loadAll();

    assert.equal(await h.host.runAction('broken/half'), false);
    assert.equal(await h.host.runAction('healthy/ok'), true);
    assert.deepEqual(h.errors, ['Activation failed: boom']);
  } finally {
    h.cleanup();
  }
});

test('rejects entry files outside the plugin directory', async () => {
  const h = createHarness({ escape: enabled() });
  try {
    writePlugin(h.root, 'escape', {
      'plugin.json': JSON.stringify({ main: '../elsewhere.cjs' }),
    });

    await h.host.loadAll();

    assert.equal(h.errors.length, 1);
    assert.match(h.errors[0]!, /must stay inside/);
  } finally {
    h.cleanup();
  }
});

test('mpv listener errors are contained and unload detaches listeners and calls deactivate', async () => {
  const h = createHarness({ watcher: enabled() });
  const client = createFakeMpvClient();
  try {
    writePlugin(h.root, 'watcher', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => {
        ctx.mpv.on('pause-change', ({ paused }) => {
          ctx.notify('paused=' + paused);
          throw new Error('listener bug');
        });
        return () => ctx.notify('deactivated');
      };`,
    });
    h.host.attachMpvClient(client);
    await h.host.loadAll();

    client.emit('pause-change', { paused: true });
    assert.deepEqual(h.notifications, ['paused=true']);
    assert.deepEqual(h.errors, ['mpv pause-change listener failed: listener bug']);

    await h.host.unloadAll();
    client.emit('pause-change', { paused: false });
    assert.deepEqual(h.notifications, ['paused=true', 'deactivated']);
  } finally {
    h.cleanup();
  }
});

test('unload gives up on a deactivate that never settles and still disposes the plugin', async () => {
  const h = createHarness({ stuck: enabled() }, { deactivateTimeoutMs: 20 });
  const client = createFakeMpvClient();
  try {
    writePlugin(h.root, 'stuck', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => {
        ctx.mpv.on('pause-change', () => ctx.notify('still listening'));
        return () => new Promise(() => {});
      };`,
    });
    h.host.attachMpvClient(client);
    await h.host.loadAll();

    await h.host.unloadAll();
    client.emit('pause-change', { paused: true });

    assert.deepEqual(h.notifications, []);
    assert.deepEqual(h.errors, ['deactivate failed: timed out after 20 ms']);
  } finally {
    h.cleanup();
  }
});

test('events follow a replacement mpv client', async () => {
  const h = createHarness({ watcher: enabled() });
  const first = createFakeMpvClient();
  const second = createFakeMpvClient();
  try {
    writePlugin(h.root, 'watcher', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => {
        ctx.mpv.on('media-path-change', ({ path }) => ctx.notify(path));
      };`,
    });
    h.host.attachMpvClient(first);
    await h.host.loadAll();
    h.host.attachMpvClient(second);

    first.emit('media-path-change', { path: 'stale.mkv' });
    second.emit('media-path-change', { path: 'current.mkv' });

    assert.deepEqual(h.notifications, ['current.mkv']);
  } finally {
    h.cleanup();
  }
});

test('unload skips modules whose manifest is still being read and prevents later loads', async () => {
  const h = createHarness({ late: enabled() });
  const client = createFakeMpvClient();
  try {
    writePlugin(h.root, 'late', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `throw new Error('module must not execute after unload starts');`,
    });
    h.host.attachMpvClient(client);

    // The plugin is still reading its manifest when unload starts.
    const loading = h.host.loadAll();
    await h.host.unloadAll();
    await loading;
    await h.host.loadAll();
    client.emit('pause-change', { paused: true });

    assert.deepEqual(h.notifications, []);
    assert.deepEqual(h.errors, []);
  } finally {
    h.cleanup();
  }
});

test('a plugin that finishes activating after unload gave up unloads itself', async () => {
  const started = deferred();
  const resume = deferred();
  const cleaned = deferred();
  let deactivated = 0;
  const h = createHarness(
    { slow: enabled() },
    {
      deactivateTimeoutMs: 20,
      loadModule: () => ({
        activate: async (ctx) => {
          started.resolve();
          await resume.promise;
          ctx.mpv.on('pause-change', () => ctx.notify('still listening'));
          return async () => {
            deactivated++;
            await Promise.resolve();
            cleaned.resolve();
          };
        },
      }),
    },
  );
  const client = createFakeMpvClient();
  try {
    writePlugin(h.root, 'slow', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
    });
    h.host.attachMpvClient(client);

    const loading = h.host.loadAll();
    await started.promise;
    await h.host.unloadAll();
    resume.resolve();
    await loading;
    await cleaned.promise;
    await h.host.unloadAll();
    client.emit('pause-change', { paused: true });

    assert.equal(deactivated, 1);
    assert.deepEqual(h.notifications, []);
    assert.deepEqual(h.errors, []);
  } finally {
    h.cleanup();
  }
});

test('unload awaits asynchronous cleanup once when activation finishes during unload', async () => {
  const started = deferred();
  const resume = deferred();
  const deactivating = deferred();
  const finish = deferred();
  let deactivated = 0;
  const h = createHarness(
    { slow: enabled() },
    {
      loadModule: () => ({
        activate: async (ctx) => {
          started.resolve();
          await resume.promise;
          return async () => {
            deactivated++;
            deactivating.resolve();
            await finish.promise;
            await ctx.anki.invoke('sync');
            ctx.notify('cleaned');
          };
        },
      }),
    },
  );
  try {
    writePlugin(h.root, 'slow', { 'plugin.json': '{"main":"main.cjs"}' });
    const loading = h.host.loadAll();
    await started.promise;
    let unloaded = false;
    const unloading = Promise.all([h.host.unloadAll(), h.host.unloadAll()]).then(() => {
      unloaded = true;
    });
    resume.resolve();
    await deactivating.promise;
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(unloaded, false);
    assert.equal(deactivated, 1);
    finish.resolve();
    await unloading;
    await loading;
    assert.deepEqual(h.notifications, ['cleaned']);
    assert.deepEqual(h.errors, []);
  } finally {
    finish.resolve();
    resume.resolve();
    await h.host.unloadAll();
    h.cleanup();
  }
});

test('a timed-out plugin cannot call app APIs after disposal', async () => {
  let context: PluginContext | undefined;
  const calls: string[] = [];
  const h = createHarness(
    { stuck: enabled() },
    {
      activateTimeoutMs: 10,
      loadModule: () => ({
        activate: (ctx) => {
          context = ctx;
          return new Promise(() => {});
        },
      }),
      sendMpvCommand: () => {
        calls.push('command');
      },
      requestMpvCommand: async () => {
        calls.push('request');
      },
      requestMpvProperty: async () => {
        calls.push('property');
      },
      ensureMpvRunning: async () => {
        calls.push('start');
        return true;
      },
      invokeAnkiConnect: async () => {
        calls.push('anki');
      },
    },
  );
  try {
    writePlugin(h.root, 'stuck', { 'plugin.json': '{"main":"main.cjs"}' });
    await h.host.loadAll();
    assert.ok(context);
    const ctx = context;
    assert.equal(ctx.signal.aborted, true);
    assert.throws(() => ctx.mpv.command(['stop']), /disposed/);
    assert.throws(() => ctx.notify('late'), /disposed/);
    await assert.rejects(ctx.mpv.request(['stop']), /disposed/);
    await assert.rejects(ctx.mpv.getProperty('path'), /disposed/);
    await assert.rejects(ctx.mpv.ensureRunning(), /disposed/);
    await assert.rejects(ctx.anki.invoke('sync'), /disposed/);
    await assert.rejects(ctx.actions.run('openJimaku'), /disposed/);
    assert.deepEqual(calls, []);
    assert.deepEqual(h.notifications, []);
    assert.deepEqual(h.sessionActions, []);
  } finally {
    await h.host.unloadAll();
    h.cleanup();
  }
});

test('runAction waits for a load that is still in progress', async () => {
  const h = createHarness({ slow: enabled() });
  try {
    writePlugin(h.root, 'slow', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = async (ctx) => {
        await new Promise((resolve) => setTimeout(resolve, 20));
        ctx.actions.register('go', () => ctx.notify('went'));
      };`,
    });

    void h.host.loadAll();

    assert.equal(await h.host.runAction('slow/go'), true);
    assert.deepEqual(h.notifications, ['went']);
  } finally {
    h.cleanup();
  }
});

test('a stuck activation cannot block another plugin and times out its own registrations', async () => {
  const h = createHarness({ stuck: enabled(), healthy: enabled() }, { activateTimeoutMs: 30 });
  try {
    writePlugin(h.root, 'stuck', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => {
        ctx.actions.register('go', () => ctx.notify('stuck ran'));
        return new Promise(() => {});
      };`,
    });
    writePlugin(h.root, 'healthy', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => ctx.actions.register('go', () => ctx.notify('healthy ran'));`,
    });
    const loading = h.host.loadAll();
    assert.equal(await h.host.runAction('healthy/go'), true);
    assert.deepEqual(h.notifications, ['healthy ran']);
    assert.deepEqual(h.errors, []);
    await loading;
    assert.equal(await h.host.runAction('stuck/go'), false);
    assert.match(h.errors[0]!, /timed out/);
  } finally {
    await h.host.unloadAll();
    h.cleanup();
  }
});

test('lifetime aborts before deactivate and disposes early and late registrations once', async () => {
  const events: string[] = [];
  const h = createHarness(
    { lifetime: enabled() },
    {
      loadModule: () => ({
        activate: (ctx) => {
          const dispose = ctx.onDispose(() => events.push('early'));
          dispose();
          dispose();
          ctx.onDispose(() => events.push('cleanup'));
          ctx.signal.addEventListener('abort', () => events.push('aborted'));
          return () => {
            events.push('deactivate aborted=' + ctx.signal.aborted);
            ctx.onDispose(() => events.push('late cleanup'));
            ctx.after(0, () => events.push('leaked timer'));
            ctx.actions.register('late', () => events.push('leaked action'));
          };
        },
      }),
    },
  );
  try {
    writePlugin(h.root, 'lifetime', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
    });
    await h.host.loadAll();
    await h.host.unloadAll();
    await new Promise((resolve) => setTimeout(resolve, 10));
    assert.equal(await h.host.runAction('lifetime/late'), false);
    assert.deepEqual(events, [
      'early',
      'aborted',
      'deactivate aborted=true',
      'late cleanup',
      'cleanup',
    ]);
  } finally {
    h.cleanup();
  }
});

test('media subscriptions replay the initial state and follow repeated loads and disconnects', async () => {
  const h = createHarness({ media: enabled() });
  const client = createFakeMpvClient();
  try {
    writePlugin(h.root, 'media', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => {
        ctx.media.on('change', (media) => ctx.notify(JSON.stringify(media)));
      };`,
    });
    h.host.attachMpvClient(client);
    client.emit('media-path-change', { path: 'episode.mkv' });
    client.emit('media-title-change', { title: 'Episode 1' });
    await h.host.loadAll();
    const first = h.host.getMediaSnapshot();
    assert.ok(first);
    assert.deepEqual(h.notifications, [JSON.stringify(first)]);
    client.emit('file-loaded', undefined);
    client.emit('file-loaded', undefined);
    assert.notEqual(h.host.getMediaSnapshot()?.playbackId, first.playbackId);
    assert.equal(first.title, 'Episode 1');
    client.emit('media-path-change', { path: 'next.mkv' });
    assert.equal(h.host.getMediaSnapshot()?.title, null);
    client.emit('connection-change', { connected: false });
    assert.equal(h.host.getMediaSnapshot(), null);
    await h.host.unloadAll();
    const count = h.notifications.length;
    client.emit('media-path-change', { path: 'later.mkv' });
    assert.equal(h.notifications.length, count);
  } finally {
    await h.host.unloadAll();
    h.cleanup();
  }
});

test('incompatible API versions are rejected before the entry module executes', async () => {
  const h = createHarness({ future: enabled() });
  try {
    writePlugin(h.root, 'future', {
      'plugin.json': JSON.stringify({ apiVersion: 999, main: 'main.cjs' }),
      'main.cjs': `throw new Error('should never execute');`,
    });
    await h.host.loadAll();
    assert.match(h.errors[0]!, /requires API 999/);
    assert.equal(h.errors.length, 1);
  } finally {
    h.cleanup();
  }
});

test('a malformed manifest version cannot prevent compatible plugins from loading', async () => {
  const h = createHarness({ bad: enabled(), good: enabled() });
  try {
    writePlugin(h.root, 'bad', {
      'plugin.json': JSON.stringify({ main: 'main.cjs', apiVersion: { toString: null } }),
      'main.cjs': `throw new Error('should never execute');`,
    });
    writePlugin(h.root, 'good', {
      'plugin.json': JSON.stringify({ main: 'main.cjs', apiVersion: 1 }),
      'main.cjs': `exports.activate = (ctx) => { ctx.actions.register('go', () => ctx.notify('good')); };`,
    });
    await h.host.loadAll();
    assert.equal(await h.host.runAction('good/go'), true);
    assert.deepEqual(h.notifications, ['good']);
    assert.deepEqual(h.errors, ['plugin.json "apiVersion" must be a positive integer.']);
  } finally {
    await h.host.unloadAll();
    h.cleanup();
  }
});

test('episode-tags uses mining provenance for a kept note older than this playback', async () => {
  const calls: unknown[] = [];
  const h = createHarness(
    { tags: enabled() },
    {
      invokeAnkiConnect: async (action, params) => {
        calls.push({ action, params });
      },
    },
  );
  try {
    writePlugin(h.root, 'tags', {
      'plugin.json': JSON.stringify({ main: 'main.cjs', apiVersion: 1 }),
      'main.cjs': `module.exports = require(${JSON.stringify(path.resolve(__dirname, '../../../../assets/plugins/episode-tags/main.js'))});`,
    });
    await h.host.loadAll();
    h.host.emitAnkiEvent('card-mined', {
      noteIds: [7],
      notes: [
        {
          originalNoteId: 1234,
          noteId: 7,
          provenance: {
            minedAt: 1234,
            media: { playbackId: 1, path: 'Show - 02.mkv', title: null },
            subtitle: null,
          },
        },
      ],
    });
    await new Promise((resolve) => setImmediate(resolve));
    assert.deepEqual(calls, [
      { action: 'addTags', params: { notes: [7], tags: 'show::Show::E02' } },
    ]);
  } finally {
    await h.host.unloadAll();
    h.cleanup();
  }
});

test('each plugin gets its own data folder, created before activate', async () => {
  const h = createHarness({ keeper: enabled() });
  try {
    writePlugin(h.root, 'keeper', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `const fs = require('node:fs');
      exports.activate = (ctx) => {
        fs.writeFileSync(require('node:path').join(ctx.dataDir, 'state.json'), '{}');
        ctx.notify(ctx.dataDir);
      };`,
    });

    await h.host.loadAll();

    const expected = path.join(h.dataRoot, 'keeper');
    assert.deepEqual(h.notifications, [expected]);
    assert.equal(existsSync(path.join(expected, 'state.json')), true);
    assert.deepEqual(h.errors, []);
  } finally {
    h.cleanup();
  }
});

test('anki card-mined reaches listeners until unload, and invoke forwards to AnkiConnect', async () => {
  const invoked: unknown[][] = [];
  const h = createHarness(
    { syncer: enabled() },
    {
      invokeAnkiConnect: async (...args) => {
        invoked.push(args);
        return 'synced';
      },
    },
  );
  try {
    writePlugin(h.root, 'syncer', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => {
        ctx.anki.on('card-mined', async ({ noteIds }) => {
          ctx.notify(noteIds.join(',') + ' ' + (await ctx.anki.invoke('sync', undefined, { timeoutMs: 5 })));
        });
      };`,
    });
    await h.host.loadAll();

    h.host.emitAnkiEvent('card-mined', { noteIds: [1, 2], notes: [] });
    await new Promise((resolve) => setImmediate(resolve));
    assert.deepEqual(h.notifications, ['1,2 synced']);
    assert.deepEqual(invoked, [['sync', {}, { timeoutMs: 5 }]]);

    await h.host.unloadAll();
    h.host.emitAnkiEvent('card-mined', { noteIds: [3], notes: [] });
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(h.notifications.length, 1);
  } finally {
    h.cleanup();
  }
});

test('actions.run forwards built-in session actions', async () => {
  const h = createHarness({ caller: enabled() });
  try {
    writePlugin(h.root, 'caller', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = async (ctx) => {
        await ctx.actions.run('openJimaku');
        await ctx.actions.run('cycleRuntimeOption', { runtimeOptionId: 'x', direction: -1 });
      };`,
    });

    await h.host.loadAll();

    assert.deepEqual(h.sessionActions, [
      { actionId: 'openJimaku' },
      { actionId: 'cycleRuntimeOption', payload: { runtimeOptionId: 'x', direction: -1 } },
    ]);
  } finally {
    h.cleanup();
  }
});

test('openWindow resolves plugin files, reuses open windows, and closes them on unload', async () => {
  const h = createHarness({ viewer: enabled() });
  try {
    writePlugin(h.root, 'viewer', {
      'plugin.json': JSON.stringify({ main: 'main.cjs' }),
      'main.cjs': `exports.activate = (ctx) => {
        ctx.actions.register('open', () => ctx.ui.openWindow({ html: 'ui/index.html', preload: 'preload.js', width: 400 }));
        ctx.actions.register('escape', () => ctx.ui.openWindow({ html: '../other/index.html' }));
      };`,
    });
    await h.host.loadAll();

    await h.host.runAction('viewer/open');
    await h.host.runAction('viewer/open');
    const dir = path.join(h.root, 'viewer');
    assert.deepEqual(h.windowRequests[0], {
      key: `viewer:${path.join(dir, 'ui', 'index.html')}`,
      htmlPath: path.join(dir, 'ui', 'index.html'),
      preloadPath: path.join(dir, 'preload.js'),
      title: 'viewer',
      width: 400,
      height: 720,
    });
    assert.equal(h.windows.size, 1);

    await h.host.runAction('viewer/escape');
    assert.equal(h.windowRequests.length, 2);
    assert.match(h.errors[0] ?? '', /must stay inside/);

    const [window] = [...h.windows.values()];
    await h.host.unloadAll();
    assert.equal(window?.destroyed, true);
  } finally {
    h.cleanup();
  }
});
