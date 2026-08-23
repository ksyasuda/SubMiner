import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

const source = readFileSync('scripts/prepare-build-assets.mjs', 'utf8');

test('macOS helper build creates dist scripts directory before swiftc output', () => {
  const buildFunctionIndex = source.indexOf('function buildMacosHelper()');
  assert.notEqual(buildFunctionIndex, -1);

  const swiftcIndex = source.indexOf("'swiftc'", buildFunctionIndex);
  assert.notEqual(swiftcIndex, -1);

  const ensureDirIndex = source.lastIndexOf('ensureDir(scriptsOutputDir)', swiftcIndex);

  assert.ok(
    ensureDirIndex > buildFunctionIndex,
    'buildMacosHelper must create dist/scripts before swiftc writes the helper binary',
  );
});

// Regression guard for #213: an untargeted swiftc stamps the build machine's OS
// version as the helper's minimum, so released builds refuse to load on older macOS.
test('macOS helper is compiled with an explicit deployment target', () => {
  assert.match(source, /-target/);
  assert.match(source, /apple-macos\$\{MACOS_HELPER_DEPLOYMENT_TARGET\}/);
});
