import assert from 'node:assert/strict';
import { mkdtemp, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { importYomitanDictionaryFromZip } from './yomitan-parser-runtime';
import { createDeps } from './yomitan-scan-test-harness';

// Hachidori deps whose sharing status reports `host`, recording every archive the
// settings automation imports (locally or forwarded over the link).
function hachidoriDeps(host: Record<string, unknown> | null) {
  const imported: string[] = [];
  const runScript = async (script: string): Promise<unknown> => {
    if (script.includes('hd_sharing_status')) {
      return {
        ok: true,
        sharing: {
          client: host
            ? { linked: true, connected: true, address: 'ws://desktop:8771/link', host }
            : { linked: false },
        },
      };
    }
    const archiveUrl = /importDictionaryArchiveUrl\(\s*("[^"]+")/.exec(script)?.[1];
    if (archiveUrl) imported.push(await (await fetch(JSON.parse(archiveUrl))).text());
    return true;
  };
  const settingsWindow = {
    isDestroyed: () => false,
    destroy: () => undefined,
    webContents: { executeJavaScript: runScript },
  };
  return {
    imported,
    deps: {
      ...createDeps(runScript, { createYomitanExtensionWindow: async () => settingsWindow }),
      getYomitanExt: () => ({
        id: 'hachi',
        name: 'Hachidori',
        version: '1',
        path: '',
        url: '',
        manifest: {},
      }),
    },
  };
}

async function writeArchive(): Promise<string> {
  const zipPath = path.join(await mkdtemp(path.join(os.tmpdir(), 'hachi-import-')), 'merged.zip');
  await writeFile(zipPath, 'PK-test');
  return zipPath;
}

for (const [name, host] of [
  ['an unlinked Hachidori imports the ZIP itself', null],
  [
    'a linked host that accepts imports gets the ZIP over the link',
    { name: 'Helium', dictionaryCount: 8, capabilities: ['linked-import-v1'] },
  ],
] as const) {
  test(name, async () => {
    const { deps, imported } = hachidoriDeps(host);
    assert.equal(
      await importYomitanDictionaryFromZip(await writeArchive(), deps, { error: assert.fail }),
      true,
    );
    assert.deepEqual(imported, ['PK-test']);
  });
}

test('a linked host without linked imports fails with the update or manual-import fix', async () => {
  const { deps, imported } = hachidoriDeps({ name: 'Chrome', dictionaryCount: 8 });
  const zipPath = await writeArchive();
  const errors: string[] = [];
  assert.equal(
    await importYomitanDictionaryFromZip(zipPath, deps, {
      error: (...parts: unknown[]) => errors.push(parts.join(' ')),
    }),
    false,
  );
  assert.deepEqual(imported, []);
  assert.match(
    errors[0] ?? '',
    /Update its Hachidori to 0\.2\.3 or later, or import .*merged\.zip/,
  );
});
