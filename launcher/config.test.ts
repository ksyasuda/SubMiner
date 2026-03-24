import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import path from 'node:path';

test('launcher root help lists subcommands', () => {
  const output = execFileSync('bun', ['run', path.join(process.cwd(), 'launcher/main.ts'), '-h'], {
    encoding: 'utf8',
  });

  assert.match(output, /Commands:/);
  assert.match(output, /jellyfin\|jf/);
  assert.match(output, /doctor/);
  assert.match(output, /config/);
  assert.match(output, /mpv/);
  assert.match(output, /dictionary\|dict/);
  assert.match(output, /texthooker/);
  assert.match(output, /app\|bin/);
});
