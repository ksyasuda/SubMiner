import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type { AnimeBridgeClient } from '../../anime-bridge/bridge-client';
import type { AnimeBrowserBridgeState } from '../../types/anime-browser';
import type { BridgeInstall, StagedBridgeUpdate } from './anime-bridge-installer';
import { createAnimeBrowserRuntime, type AnimeBrowserRuntimeDeps } from './anime-browser-runtime';

// The installer never fills `updateAvailable`; the runtime asks upstream later.
const OLD: BridgeInstall = {
  javaPath: '/managed/jre/bin/java',
  jarPath: '/managed/MExtensionServer-v1.0.5.0.jar',
  dir: '/managed',
  origin: 'managed',
  version: 'v1.0.5.0',
  updateAvailable: null,
};

const NEW: BridgeInstall = { ...OLD, version: 'v1.0.6.0' };

/** Upstream's newest release, as the update check would report it. */
const LATEST = 'v1.0.6.0';

const tick = () => new Promise((resolve) => setTimeout(resolve, 10));

async function setup(overrides: Partial<AnimeBrowserRuntimeDeps> = {}) {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-anime-bridge-update-'));
  const states: AnimeBrowserBridgeState[] = [];
  const stopped: number[] = [];
  let started = 0;
  let current = OLD;
  const runtime = createAnimeBrowserRuntime({
    extensionsDir: () => dir,
    repos: () => [],
    setRepos: () => undefined,
    preferencesFile: path.join(dir, 'preferences.json'),
    ensureBinaries: async () => current,
    checkBridgeUpdate: async (install) => (install.version === LATEST ? null : LATEST),
    stageBridgeUpdate: async (onProgress) => {
      onProgress({ stage: 'downloading', progress: 0.5 });
      return {
        version: NEW.version!,
        commit: async () => {
          current = NEW;
          return NEW;
        },
      };
    },
    sendMpvCommand: () => undefined,
    ensureMpvConnected: async () => true,
    onBridgeState: (state) => states.push(state),
    log: () => undefined,
    startSidecar: async () => {
      const id = ++started;
      return {
        client: { listAnimeSources: async () => [] } as unknown as AnimeBridgeClient,
        baseUrl: `http://127.0.0.1:${id}`,
        port: id,
        stop: async () => {
          stopped.push(id);
        },
        onExit: () => undefined,
      };
    },
    startStreamStripProxy: async () => ({
      origin: 'http://127.0.0.1:9',
      port: 9,
      close: async () => undefined,
    }),
    ...overrides,
  });
  return { runtime, states, stopped, started: () => started };
}

test('the bridge state says where the running bridge came from, then learns about an update', async () => {
  const { runtime, states } = await setup();
  const state = await runtime.ensureBridge();
  assert.equal(state.stage, 'ready');
  assert.equal(state.install?.origin, 'managed');
  assert.equal(state.install?.version, 'v1.0.5.0');
  assert.equal(state.install?.dir, '/managed');

  // The upstream check is kicked off after ready, off the start path, and
  // re-broadcasts the ready state once it answers.
  await tick();
  const latest = runtime.getSnapshot().bridge;
  assert.equal(latest.stage, 'ready');
  assert.equal(latest.install?.updateAvailable, LATEST);
  assert.equal(states.at(-1)?.install?.updateAvailable, LATEST);
});

test('a failed update check is logged and leaves the bridge ready', async () => {
  const logged: string[] = [];
  const { runtime } = await setup({
    checkBridgeUpdate: async () => {
      throw new Error('rate limited');
    },
    log: (message) => logged.push(message),
  });
  await runtime.ensureBridge();
  await tick();
  const state = runtime.getSnapshot().bridge;
  assert.equal(state.stage, 'ready');
  assert.equal(state.install?.updateAvailable, null);
  assert.ok(logged.some((line) => /update check failed: rate limited/.test(line)));
});

test('a system install is never asked about updates', async () => {
  let asked = 0;
  const { runtime } = await setup({
    ensureBinaries: async () => ({ ...OLD, origin: 'system' }),
    checkBridgeUpdate: async () => {
      asked += 1;
      return LATEST;
    },
  });
  await runtime.ensureBridge();
  await tick();
  assert.equal(asked, 0);
  assert.equal(runtime.getSnapshot().bridge.install?.updateAvailable, null);
});

test('updateBridge stages, stops the old bridge, and restarts on the new install', async () => {
  const { runtime, states, stopped, started } = await setup();
  await runtime.ensureBridge();

  await tick();
  assert.equal(runtime.getSnapshot().bridge.install?.updateAvailable, LATEST);

  const state = await runtime.updateBridge();

  assert.deepEqual(stopped, [1]);
  assert.equal(started(), 2);
  assert.equal(state.stage, 'ready');
  assert.equal(state.install?.version, 'v1.0.6.0');
  // Progress from the staging download reached the UI, still against the old install.
  const downloading = states.find((candidate) => candidate.stage === 'downloading');
  assert.equal(downloading?.progress, 0.5);
  assert.equal(downloading?.install?.version, 'v1.0.5.0');
  // The post-restart check finds nothing newer, so the button goes away.
  await tick();
  assert.equal(runtime.getSnapshot().bridge.install?.updateAvailable, null);
});

test('a failed download leaves the running bridge as it was', async () => {
  const { runtime, stopped, started } = await setup({
    stageBridgeUpdate: async () => {
      throw new Error('offline');
    },
  });
  await runtime.ensureBridge();

  const state = await runtime.updateBridge();

  assert.deepEqual(stopped, []);
  assert.equal(started(), 1);
  assert.equal(state.stage, 'ready');
  assert.match(state.message ?? '', /Bridge update failed: offline/);
});

test('updateBridge refuses to touch a system install', async () => {
  const system: BridgeInstall = {
    ...OLD,
    origin: 'system',
    dir: '/usr/share/mangatan/extension_server',
  };
  const { runtime } = await setup({ ensureBinaries: async () => system });
  await runtime.ensureBridge();
  await assert.rejects(runtime.updateBridge(), /managed outside SubMiner/);
});

test('requests that arrive mid-update wait for the new bridge instead of starting the old one', async () => {
  let releaseCommit: () => void = () => undefined;
  const commitGate = new Promise<void>((resolve) => {
    releaseCommit = resolve;
  });
  let current = OLD;
  const staged: StagedBridgeUpdate = {
    version: 'v1.0.6.0',
    commit: async () => {
      await commitGate;
      current = NEW;
      return NEW;
    },
  };
  const { runtime, started } = await setup({
    ensureBinaries: async () => current,
    stageBridgeUpdate: async () => staged,
  });
  await runtime.ensureBridge();

  const update = runtime.updateBridge();
  // Let the update get past stopping the old bridge and into the gated commit.
  await new Promise((resolve) => setTimeout(resolve, 10));
  let ensured: AnimeBrowserBridgeState | null = null;
  const waiting = runtime.ensureBridge().then((state) => {
    ensured = state;
  });
  await new Promise((resolve) => setTimeout(resolve, 10));
  assert.equal(ensured, null);
  assert.equal(started(), 1);

  releaseCommit();
  await update;
  await waiting;
  assert.equal(started(), 2);
  assert.equal(ensured!.install?.version, 'v1.0.6.0');
});

test('an update waits for an in-flight start and stops its sidecar before commit', async () => {
  let releaseFirstStart: () => void = () => undefined;
  const firstStartGate = new Promise<void>((resolve) => {
    releaseFirstStart = resolve;
  });
  let markFirstStartEntered: () => void = () => undefined;
  const firstStartEntered = new Promise<void>((resolve) => {
    markFirstStartEntered = resolve;
  });
  const events: string[] = [];
  let startCount = 0;
  let current = OLD;
  const { runtime } = await setup({
    ensureBinaries: async () => current,
    stageBridgeUpdate: async () => ({
      version: LATEST,
      commit: async () => {
        events.push('commit');
        current = NEW;
        return NEW;
      },
    }),
    startSidecar: async () => {
      const id = ++startCount;
      events.push(`start:${id}`);
      if (id === 1) {
        markFirstStartEntered();
        await firstStartGate;
      }
      return {
        client: { listAnimeSources: async () => [] } as unknown as AnimeBridgeClient,
        baseUrl: `http://127.0.0.1:${id}`,
        port: id,
        stop: async () => {
          events.push(`stop:${id}`);
        },
        onExit: () => undefined,
      };
    },
  });

  const firstStart = runtime.ensureBridge();
  await firstStartEntered;
  const update = runtime.updateBridge();
  await tick();
  assert.deepEqual(events, ['start:1']);

  releaseFirstStart();
  await Promise.all([firstStart, update]);

  assert.deepEqual(events, ['start:1', 'stop:1', 'commit', 'start:2']);
  assert.equal(runtime.getSnapshot().bridge.install?.version, NEW.version);
});
