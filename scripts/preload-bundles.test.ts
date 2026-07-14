import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

test('build:syncui bundles the sandboxed preload and keeps Electron external', () => {
  const packageJson = JSON.parse(
    fs.readFileSync(path.join(import.meta.dir, '..', 'package.json'), 'utf8'),
  ) as { scripts: Record<string, string> };
  const command = packageJson.scripts['build:syncui'] ?? '';

  assert.match(command, /src\/preload-syncui\.ts/);
  assert.match(command, /--bundle/);
  assert.match(command, /--external:electron/);
  assert.match(command, /--outfile=dist\/preload-syncui\.js/);
});
