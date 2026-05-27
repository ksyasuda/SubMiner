import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

const source = readFileSync('scripts/prepare-build-assets.mjs', 'utf8');

test('macOS helper build creates dist scripts directory before swiftc output', () => {
  const buildFunctionIndex = source.indexOf('function buildMacosHelper()');
  assert.notEqual(buildFunctionIndex, -1);

  const swiftcIndex = source.indexOf("execFileSync('swiftc'", buildFunctionIndex);
  assert.notEqual(swiftcIndex, -1);

  const ensureDirIndex = source.lastIndexOf('ensureDir(scriptsOutputDir)', swiftcIndex);

  assert.ok(
    ensureDirIndex > buildFunctionIndex,
    'buildMacosHelper must create dist/scripts before swiftc writes the helper binary',
  );
});
