import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, writeFileSync, statSync, rmSync, createReadStream } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { createPackageFromStreams } from '@electron/asar';
import { FileMatcher, getFileMatchers } from 'app-builder-lib/out/fileMatcher';
import config from '../package.json';
import { listAppFiles, listFiles, compareSizes, verifyAppPath } from './package-audit.cjs';

test('platform packaging preserves the runtime allowlist after builder normalizes global filters', () => {
  const root = process.cwd();
  const fileStat = statSync('package.json');
  for (const platform of ['linux', 'mac', 'win'] as const) {
    const matchers = getFileMatchers(
      { files: [{ filter: config.build.files }] },
      'files',
      '/tmp/subminer-filter-output',
      {
        defaultSrc: root,
        globalOutDir: path.join(root, 'release'),
        customBuildOptions: { files: config.build[platform].files },
        macroExpander: (value) => value.replaceAll('${arch}', 'x64'),
      },
    );
    assert(matchers);
    // This is builder's default for an exclusion-only platform matcher.
    for (const matcher of matchers) {
      if (matcher.containsOnlyIgnore()) matcher.prependPattern('**/*');
    }
    const included = (name: string) =>
      matchers.some((matcher) => matcher.createFilter()(path.join(root, name), fileStat));
    for (const name of [
      'dist/main-entry.js',
      'dist/fonts/MPLUS1[wght].ttf',
      'stats/dist/index.html',
      'vendor/texthooker-ui/docs/index.html',
      'package.json',
    ]) {
      assert(included(name), `${platform} must ship ${name}`);
    }
    for (const name of [
      '.agents/skills/test.md',
      'src/main.ts',
      'scripts/build-yomitan.mjs',
      'docs-site/index.md',
      'dist/main.js.map',
      'dist/main.test.js',
      'dist/nested/source.ts',
      'dist/nested/__tests__/helper.js',
      'stats/dist/nested/fixtures/data.json',
      'vendor/texthooker-ui/docs/nested/component.tsx',
      'dist/launcher/subminer',
      'dist/settings/fonts/MPLUS1[wght].ttf',
      'vendor/subminer-yomitan/ext/manifest.json',
    ]) {
      assert(!included(name), `${platform} must exclude ${name}`);
    }
  }
});

test('dependency filters keep only the target Windows Koffi binary', () => {
  const root = process.cwd();
  for (const arch of ['x64', 'arm64']) {
    for (const platform of ['linux', 'mac', 'win'] as const) {
      const patterns = [
        '**/*',
        ...config.build.files.filter((name) => name.startsWith('!')),
        ...config.build[platform].files.filter((name) => name.startsWith('!')),
      ];
      const filter = new FileMatcher(
        root,
        '/tmp/subminer-filter-output',
        (value) => value.replaceAll('${arch}', arch),
        patterns,
      ).createFilter();
      const included = (name: string) =>
        filter(path.join(root, 'node_modules', name), statSync('package.json'));
      assert(included('@libsql/win32-x64-msvc/index.node'));
      assert(!included('axios/dist/axios.js.map'));
      assert(!included('koffi/src/koffi/src/ffi.c'));
      assert(!included('agent-base/src/index.ts'));
      assert(!included('@discordjs/rest/dist/index.d.mts'));
      assert(!included('example/lib/tests/helper.js'));
      for (const target of [
        'win32_x64',
        'win32_arm64',
        'linux_x64',
        'darwin_arm64',
        'openbsd_x64',
      ]) {
        assert.equal(
          included(`koffi/build/koffi/${target}/koffi.node`),
          platform === 'win' && target === `win32_${arch}`,
          `${platform}/${arch}: ${target}`,
        );
      }
      assert.equal(included('koffi/index.js'), platform === 'win');
      assert.equal(included('koffi/LICENSE.txt'), platform === 'win');
    }
  }
});

test('content audit rejects development files beneath approved roots', () => {
  for (const root of ['dist', 'stats/dist', 'vendor/texthooker-ui/docs', 'node_modules/example']) {
    for (const suffix of [
      'nested/source.ts',
      'nested/component.tsx',
      'nested/types.d.mts',
      'nested/source.cts',
      'nested/__tests__/helper.js',
      'nested/tests/helper.js',
      'nested/test/helper.js',
      'nested/__fixtures__/data.json',
      'nested/fixtures/data.json',
      'nested/fixture/data.json',
      'nested/component.spec.js',
      'nested/component.test.cjs',
    ]) {
      assert.throws(() => verifyAppPath(`${root}/${suffix}`, 'linux', 'x64'), /Packaged/);
    }
    for (const suffix of ['nested/runtime.js', 'nested/style.css', 'nested/data.json']) {
      assert.doesNotThrow(() => verifyAppPath(`${root}/${suffix}`, 'linux', 'x64'));
    }
  }
});

test('archive inventory handles native files without counting them twice on disk', async () => {
  const root = mkdtempSync(path.join(tmpdir(), 'subminer-audit-'));
  try {
    const input = path.join(root, 'input');
    const output = path.join(root, 'output');
    mkdirSync(input);
    mkdirSync(output);
    writeFileSync(path.join(input, 'main.js'), 'hello');
    writeFileSync(path.join(input, 'native.node'), 'native');
    const archive = path.join(output, 'app.asar');
    await createPackageFromStreams(
      archive,
      ['main.js', 'native.node'].map((name) => ({
        path: name,
        type: 'file',
        unpacked: name.endsWith('.node'),
        stat: statSync(path.join(input, name)),
        streamGenerator: () => createReadStream(path.join(input, name)),
      })),
    );
    assert.deepEqual(listAppFiles(archive), [
      { path: 'main.js', bytes: 5 },
      { path: 'native.node', bytes: 6 },
    ]);
    assert.equal(
      listFiles(output).reduce((sum: number, entry: { bytes: number }) => sum + entry.bytes, 0),
      statSync(archive).size + 6,
    );
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('size comparison tolerates older reports without artifact measurements', () => {
  const previous = { version: '0.19.6', platform: 'linux', arch: 'x64', unpackedBytes: 100 };
  const current = { ...previous, unpackedBytes: 80, artifacts: [{ kind: 'AppImage', bytes: 40 }] };
  assert.deepEqual(compareSizes(current, previous), {
    version: '0.19.6',
    unpackedDeltaBytes: -20,
    artifacts: [],
  });
  assert.deepEqual(
    compareSizes(current, { ...previous, artifacts: [null, { kind: 'AppImage', bytes: 50 }] })
      .artifacts,
    [{ kind: 'AppImage', deltaBytes: -10 }],
  );
  assert.throws(
    () => compareSizes(current, { ...previous, unpackedBytes: 'unknown' }),
    /Invalid previous size report/,
  );
});
