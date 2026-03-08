import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveProxyCommandOsdRuntime } from './mpv-proxy-osd';

function createClient() {
  return {
    connected: true,
    requestProperty: async (name: string) => {
      if (name === 'sid') return 3;
      if (name === 'secondary-sid') return 8;
      if (name === 'track-list') {
        return [
          { id: 3, type: 'sub', title: 'Japanese', selected: true, external: false },
          { id: 8, type: 'sub', title: 'English Commentary', external: true },
        ];
      }
      return null;
    },
  };
}

test('resolveProxyCommandOsdRuntime formats the active primary subtitle track label', async () => {
  const result = await resolveProxyCommandOsdRuntime(['cycle', 'sid'], () => createClient());
  assert.equal(result, 'Subtitle track: Internal #3 - Japanese (active)');
});

test('resolveProxyCommandOsdRuntime formats the active secondary subtitle track label', async () => {
  const result = await resolveProxyCommandOsdRuntime(
    ['set_property', 'secondary-sid', 'auto'],
    () => createClient(),
  );
  assert.equal(result, 'Secondary subtitle track: External #8 - English Commentary');
});
