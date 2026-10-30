import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { copyFileSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

test('plugin declarations work without SubMiner sources', { timeout: 30_000 }, () => {
  const root = mkdtempSync(path.join(os.tmpdir(), 'subminer-plugin-types-'));
  try {
    copyFileSync(path.join(__dirname, 'plugin-api.ts'), path.join(root, 'subminer.d.ts'));
    symlinkSync(
      path.resolve(__dirname, '../node_modules'),
      path.join(root, 'node_modules'),
      'junction',
    );
    writeFileSync(
      path.join(root, 'tsconfig.json'),
      JSON.stringify({
        compilerOptions: {
          strict: true,
          noEmit: true,
          module: 'Node16',
          target: 'ES2022',
          types: ['node'],
        },
        files: ['plugin.ts', 'subminer.d.ts'],
      }),
    );
    writeFileSync(
      path.join(root, 'plugin.ts'),
      `
      import type { PluginContext, PluginMediaSnapshot } from './subminer';
      export function activate(ctx: PluginContext) {
        const version: 1 = ctx.apiVersion;
        const initial: PluginMediaSnapshot | null = ctx.media.current();
        ctx.onDispose(ctx.media.on('change', media => ctx.log.info(String(media?.playbackId))));
        ctx.anki.on('card-mined', ({ notes }) => {
          for (const { originalNoteId, noteId, provenance } of notes) {
            ctx.log.info(String({ originalNoteId, noteId, source: provenance?.media?.path }));
          }
        });
        ctx.actions.register('pause', async () => {
          await ctx.mpv.request(['set_property', 'pause', true]);
          await ctx.actions.run('openJimaku');
        });
        ctx.log.info(String({ version, initial, aborted: ctx.signal.aborted }));
      }
    `,
    );
    const output = execFileSync(
      process.execPath,
      [require.resolve('typescript/bin/tsc'), '-p', root],
      {
        encoding: 'utf8',
      },
    );
    assert.equal(output, '');
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
