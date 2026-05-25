import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

test('delete confirmation dialog swallows Escape before closing', () => {
  const source = fs.readFileSync(
    path.join(process.cwd(), 'stats/src/components/layout/DeleteConfirmDialog.tsx'),
    'utf8',
  );
  const handlerBlock = source.match(
    /const onKeyDown = \(event: KeyboardEvent\) => \{(?<body>[\s\S]*?)\n    \};/,
  )?.groups?.body;

  assert.ok(handlerBlock);
  assert.match(handlerBlock, /event\.preventDefault\(\);/);
  assert.match(handlerBlock, /event\.stopPropagation\(\);/);
  assert.match(handlerBlock, /event\.stopImmediatePropagation\(\);/);
  assert.ok(
    handlerBlock.indexOf('event.stopPropagation();') < handlerBlock.indexOf('finish(false);'),
  );
});
