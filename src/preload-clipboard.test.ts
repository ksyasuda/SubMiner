import assert from 'node:assert/strict';
import test from 'node:test';
import { runInNewContext } from 'node:vm';
import { build } from 'esbuild';

test('sidebar clipboard bridge writes exact text without renderer focus and rejects non-text input', async () => {
  const result = await build({
    entryPoints: ['src/preload.ts'],
    bundle: true,
    platform: 'node',
    format: 'cjs',
    external: ['electron'],
    write: false,
  });
  const output = result.outputFiles[0];
  assert.ok(output);
  const writes: string[] = [];
  let exposed: unknown;
  runInNewContext(output.text, {
    process: { argv: [] },
    require: (name: string) => {
      assert.equal(name, 'electron');
      return {
        ipcRenderer: { on: () => {} },
        clipboard: { writeText: (text: string) => writes.push(text) },
        contextBridge: {
          exposeInMainWorld: (_name: string, api: unknown) => {
            exposed = api;
          },
        },
      };
    },
  });
  assert.ok(
    typeof exposed === 'object' && exposed !== null && 'copySubtitleSidebarSelection' in exposed,
  );
  const copy = exposed.copySubtitleSidebarSelection;
  if (typeof copy !== 'function') throw new Error('Missing clipboard bridge');
  const text = '最初の台詞\n二行目\n\n同じ台詞';
  await copy(text);
  assert.deepEqual(writes, [text]);
  for (const value of [null, undefined, 42, { text }, ['台詞']]) {
    await assert.rejects(async () => copy(value), /Subtitle selection must be text/);
  }
  assert.deepEqual(writes, [text]);
});
