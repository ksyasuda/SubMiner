import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildApplyJellyfinMpvDefaultsMainDepsHandler,
  createBuildGetDefaultSocketPathMainDepsHandler,
} from './mpv-jellyfin-defaults-main-deps';

test('apply jellyfin mpv defaults main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildApplyJellyfinMpvDefaultsMainDepsHandler({
    sendMpvCommandRuntime: (_client, command) => calls.push(command.join(':')),
    jellyfinLangPref: 'ja,jp',
  })();

  deps.sendMpvCommandRuntime({ connected: true, send: () => {} }, ['set_property', 'aid', 'auto']);
  assert.equal(deps.jellyfinLangPref, 'ja,jp');
  assert.deepEqual(calls, ['set_property:aid:auto']);
});

test('get default socket path main deps builder maps platform', () => {
  const deps = createBuildGetDefaultSocketPathMainDepsHandler({
    platform: 'darwin',
  })();
  assert.equal(deps.platform, 'darwin');
});
