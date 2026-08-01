import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

function buildScript(name: string): string {
  const packageJson = JSON.parse(
    fs.readFileSync(path.join(import.meta.dir, '..', 'package.json'), 'utf8'),
  ) as { scripts: Record<string, string> };
  return packageJson.scripts[name] ?? '';
}

test('build:syncui bundles the sandboxed preload and keeps Electron external', () => {
  const command = buildScript('build:syncui');

  assert.match(command, /src\/preload-syncui\.ts/);
  assert.match(command, /--bundle/);
  assert.match(command, /--external:electron/);
  assert.match(command, /--outfile=dist\/preload-syncui\.js/);
});

test('build:animeui bundles the sandboxed preload and keeps Electron external', () => {
  const command = buildScript('build:animeui');

  // The preload imports IPC_CHANNELS, so it must be bundled rather than
  // emitted by plain tsc with a relative runtime require.
  assert.match(command, /src\/preload-animeui\.ts/);
  assert.match(command, /--bundle/);
  assert.match(command, /--external:electron/);
  assert.match(command, /--outfile=dist\/preload-animeui\.js/);
});

test('build:animeui runs as part of the top-level build', () => {
  assert.match(buildScript('build'), /bun run build:animeui/);
});
