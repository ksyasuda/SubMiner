import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

test('settings toolbar does not expose an open file button', () => {
  const html = fs.readFileSync(path.join(process.cwd(), 'src/settings/index.html'), 'utf8');

  assert.equal(html.includes('id="openFileButton"'), false);
  assert.equal(html.includes('Open File'), false);
});
