import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

test('settings preload stays sandbox-compatible by avoiding local runtime imports', () => {
  const source = fs.readFileSync(path.join(process.cwd(), 'src', 'preload-settings.ts'), 'utf8');

  assert.doesNotMatch(source, /from\s+['"]\.\/shared\/ipc\/contracts(?:\.(?:js|ts))?['"]/);
});
