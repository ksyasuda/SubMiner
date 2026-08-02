import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
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

test('anime UI stylesheet files exist and are all staged', () => {
  const html = readFileSync('src/animeui/index.html', 'utf8');
  const stylesheets = [...html.matchAll(/<link rel="stylesheet" href="\.\/(.+?\.css)"/g)].map(
    (match) => match[1],
  );

  assert.deepEqual(stylesheets, ['style.css', 'detail.css', 'panels.css']);
  for (const stylesheet of stylesheets) {
    assert.equal(existsSync(`src/animeui/${stylesheet}`), true, `${stylesheet} must exist`);
  }
  assert.match(
    source,
    /copyAssets\(animeUiSourceDir, animeUiOutputDir, 'animeui', \[\s*'style\.css',\s*'detail\.css',\s*'panels\.css',?\s*\]\)/,
  );
});
