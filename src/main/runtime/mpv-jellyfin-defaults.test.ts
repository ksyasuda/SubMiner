import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createApplyJellyfinMpvDefaultsHandler,
  createGetDefaultSocketPathHandler,
} from './mpv-jellyfin-defaults';

test('apply jellyfin mpv defaults sends expected property commands', () => {
  const calls: string[] = [];
  const applyDefaults = createApplyJellyfinMpvDefaultsHandler({
    sendMpvCommandRuntime: (_client, command) => calls.push(command.join(':')),
    jellyfinLangPref: 'ja,jp',
  });

  applyDefaults({ connected: true, send: () => {} });
  assert.deepEqual(calls, [
    'set_property:sub-auto:fuzzy',
    'set_property:aid:auto',
    'set_property:sid:auto',
    'set_property:secondary-sid:auto',
    'set_property:secondary-sub-visibility:no',
    'set_property:alang:ja,jp',
    'set_property:slang:ja,jp',
  ]);
});

test('get default socket path returns platform specific value', () => {
  const getWindowsPath = createGetDefaultSocketPathHandler({ platform: 'win32' });
  const getUnixPath = createGetDefaultSocketPathHandler({ platform: 'darwin' });
  assert.equal(getWindowsPath(), '\\\\.\\pipe\\subminer-socket');
  assert.equal(getUnixPath(), '/tmp/subminer-socket');
});
