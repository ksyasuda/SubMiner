import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const releaseWorkflowPath = resolve(__dirname, '../.github/workflows/release.yml');
const releaseWorkflow = readFileSync(releaseWorkflowPath, 'utf8');
const makefilePath = resolve(__dirname, '../Makefile');
const makefile = readFileSync(makefilePath, 'utf8');
const packageJsonPath = resolve(__dirname, '../package.json');
const packageJson = JSON.parse(readFileSync(packageJsonPath, 'utf8')) as {
  desktopName?: string;
  productName?: string;
  scripts: Record<string, string>;
  build?: {
    afterPack?: string;
    files?: string[];
  };
};

test('publish release leaves prerelease unset so gh creates a normal release', () => {
  assert.ok(!releaseWorkflow.includes('--prerelease'));
});

test('publish release forces an existing draft tag release to become public', () => {
  assert.ok(releaseWorkflow.includes('--draft=false'));
});

test('release workflow verifies a committed changelog section before publish', () => {
  assert.match(releaseWorkflow, /bun run changelog:check/);
});

test('release workflow builds changelog artifacts when fragments are pending', () => {
  assert.match(releaseWorkflow, /Build changelog artifacts for release/);
  assert.match(releaseWorkflow, /changelog:build --version/);
});

test('release workflow verifies generated config examples before packaging artifacts', () => {
  assert.match(releaseWorkflow, /bun run verify:config-example/);
});

test('release quality gate runs the maintained source coverage lane and uploads lcov output', () => {
  assert.match(releaseWorkflow, /name: Coverage suite \(maintained source lane\)/);
  assert.match(releaseWorkflow, /run: bun run test:coverage:src/);
  assert.match(releaseWorkflow, /name: Upload coverage artifact/);
  assert.match(releaseWorkflow, /path: coverage\/test-src\/lcov\.info/);
});

test('release build jobs install and cache stats dependencies before packaging', () => {
  assert.match(releaseWorkflow, /build-linux:[\s\S]*stats\/node_modules/);
  assert.match(releaseWorkflow, /build-macos:[\s\S]*stats\/node_modules/);
  assert.match(releaseWorkflow, /build-windows:[\s\S]*stats\/node_modules/);
  assert.match(releaseWorkflow, /build-linux:[\s\S]*cd stats && bun install --frozen-lockfile/);
  assert.match(releaseWorkflow, /build-macos:[\s\S]*cd stats && bun install --frozen-lockfile/);
  assert.match(releaseWorkflow, /build-windows:[\s\S]*cd stats && bun install --frozen-lockfile/);
});

test('release workflow generates release notes from committed changelog output', () => {
  assert.match(releaseWorkflow, /bun run changelog:release-notes/);
  assert.ok(!releaseWorkflow.includes('git log --pretty=format:"- %s"'));
});

test('release workflow includes the Windows installer in checksums and uploaded assets', () => {
  assert.match(
    releaseWorkflow,
    /files=\(release\/\*\.AppImage release\/\*\.dmg release\/\*\.exe release\/\*\.zip release\/\*\.tar\.gz dist\/launcher\/subminer\)/,
  );
  assert.match(
    releaseWorkflow,
    /artifacts=\([\s\S]*release\/\*\.exe[\s\S]*release\/SHA256SUMS\.txt[\s\S]*\)/,
  );
});

test('release package scripts disable implicit electron-builder publishing', () => {
  assert.match(packageJson.scripts['build:appimage'] ?? '', /--publish never/);
  assert.match(packageJson.scripts['build:mac'] ?? '', /--publish never/);
  assert.match(packageJson.scripts['build:win'] ?? '', /--publish never/);
  assert.match(packageJson.scripts['build:win:unsigned'] ?? '', /build-win-unsigned\.mjs/);
});

test('release packaging wires a shared afterPack hook for Linux AppImage library staging', () => {
  assert.equal(packageJson.build?.afterPack, 'scripts/electron-builder-after-pack.cjs');
});

test('top-level package metadata keeps Linux Electron runtime app identity canonical', () => {
  assert.equal(packageJson.productName, 'SubMiner');
  assert.equal(packageJson.desktopName, 'SubMiner.desktop');
});

test('release packaging keeps default file inclusion and excludes large source-only trees explicitly', () => {
  const files = packageJson.build?.files ?? [];
  assert.ok(files.includes('**/*'));
  assert.ok(files.includes('!src{,/**/*}'));
  assert.ok(files.includes('!launcher{,/**/*}'));
  assert.ok(files.includes('!stats/src{,/**/*}'));
  assert.ok(files.includes('!.tmp{,/**/*}'));
  assert.ok(files.includes('!release-*{,/**/*}'));
  assert.ok(files.includes('!vendor/subminer-yomitan{,/**/*}'));
  assert.ok(files.includes('!vendor/texthooker-ui/src{,/**/*}'));
  assert.ok(files.includes('!assets{,/**/*}'));
  assert.ok(files.includes('!plugin{,/**/*}'));
  assert.ok(files.includes('!vendor/yomitan-jlpt-vocab{,/**/*}'));
  assert.ok(files.includes('!docs{,/**/*}'));
  assert.ok(files.includes('!tests{,/**/*}'));
  assert.ok(files.includes('!packaging{,/**/*}'));
  assert.ok(files.includes('!README.md'));
  assert.ok(files.includes('!CHANGELOG.md'));
  assert.ok(files.includes('!AGENTS.md'));
  assert.ok(files.includes('!CLAUDE.md'));
  assert.ok(files.includes('!stats/public{,/**/*}'));
  assert.ok(files.includes('!stats/package.json'));
  assert.ok(files.includes('!stats/tsconfig.json'));
  assert.ok(files.includes('!stats/vite.config.ts'));
  assert.ok(files.includes('!dist/**/*.map'));
  assert.ok(files.includes('!dist/**/*.test.*'));
  assert.ok(files.includes('!dist/**/__tests__{,/**/*}'));
  assert.ok(files.includes('!scripts/**/*.test.*'));
  assert.ok(files.includes('!vendor/texthooker-ui/public{,/**/*}'));
  assert.ok(files.includes('!vendor/texthooker-ui/.vscode{,/**/*}'));
  assert.ok(files.includes('!vendor/texthooker-ui/README.md'));
  assert.ok(files.includes('!vendor/texthooker-ui/package.json'));
  assert.ok(files.includes('!vendor/texthooker-ui/tsconfig*.json'));
  assert.ok(files.includes('!node_modules/@libsql/linux-x64-musl{,/**/*}'));
});

test('config example generation runs directly from source without unrelated bundle prerequisites', () => {
  assert.equal(
    packageJson.scripts['generate:config-example'],
    'bun run src/generate-config-example.ts',
  );
});

test('windows release workflow publishes unsigned artifacts directly without SignPath', () => {
  assert.match(releaseWorkflow, /Build unsigned Windows artifacts/);
  assert.match(releaseWorkflow, /run: bun run build:win:unsigned/);
  assert.match(releaseWorkflow, /name: windows/);
  assert.match(releaseWorkflow, /path: \|\n\s+release\/\*\.exe\n\s+release\/\*\.zip/);
  assert.ok(!releaseWorkflow.includes('signpath/github-action-submit-signing-request'));
  assert.ok(!releaseWorkflow.includes('SIGNPATH_'));
});

test('release workflow publishes subminer-bin to AUR from tagged release artifacts', () => {
  assert.match(releaseWorkflow, /aur-publish:/);
  assert.match(releaseWorkflow, /needs:\s*\[release\]/);
  assert.match(releaseWorkflow, /AUR_SSH_PRIVATE_KEY/);
  assert.match(releaseWorkflow, /ssh:\/\/aur@aur\.archlinux\.org\/subminer-bin\.git/);
  assert.match(releaseWorkflow, /scripts\/update-aur-package\.sh/);
  assert.match(
    releaseWorkflow,
    /cp packaging\/aur\/subminer-bin\/\.SRCINFO aur-subminer-bin\/\.SRCINFO/,
  );
  assert.match(releaseWorkflow, /version_no_v="\$\{\{ steps\.version\.outputs\.VERSION \}\}"/);
  assert.match(releaseWorkflow, /SubMiner-\$\{version_no_v\}\.AppImage/);
  assert.doesNotMatch(
    releaseWorkflow,
    /SubMiner-\$\{\{ steps\.version\.outputs\.VERSION \}\}\.AppImage/,
  );
  assert.doesNotMatch(releaseWorkflow, /Install makepkg/);
});

test('release workflow skips empty AUR sync commits', () => {
  assert.match(releaseWorkflow, /if git diff --quiet -- PKGBUILD \.SRCINFO; then/);
});

test('Makefile routes Windows install-plugin setup through bun and documents Windows builds', () => {
  assert.match(
    makefile,
    /windows\) printf '%s\\n' "\[INFO\] Windows builds run via: bun run build:win" ;;/,
  );
  assert.match(makefile, /bun \.\/scripts\/configure-plugin-binary-path\.mjs/);
});
