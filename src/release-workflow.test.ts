import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const releaseWorkflowPath = resolve(__dirname, '../.github/workflows/release.yml');
const releaseWorkflow = readFileSync(releaseWorkflowPath, 'utf8');
const docsPagesWorkflowPath = resolve(__dirname, '../.github/workflows/docs-pages.yml');
const docsPagesWorkflow = readFileSync(docsPagesWorkflowPath, 'utf8');
const makefilePath = resolve(__dirname, '../Makefile');
const makefile = readFileSync(makefilePath, 'utf8');
const packageJsonPath = resolve(__dirname, '../package.json');
const packageJson = JSON.parse(readFileSync(packageJsonPath, 'utf8')) as {
  desktopName?: string;
  productName?: string;
  scripts: Record<string, string>;
  build?: {
    afterPack?: string;
    electronUpdaterCompatibility?: string;
    files?: string[];
    artifactName?: string;
    dmg?: {
      artifactName?: string;
    };
    extraResources?: Array<{
      from?: string;
      to?: string;
    }>;
    mac?: {
      artifactName?: string;
    };
    nsis?: {
      artifactName?: string;
    };
    publish?: Array<{
      provider?: string;
      owner?: string;
      repo?: string;
    }>;
    win?: {
      artifactName?: string;
    };
  };
};

test('publish release leaves prerelease unset so gh creates a normal release', () => {
  assert.ok(!releaseWorkflow.includes('--prerelease'));
});

test('stable release workflow excludes prerelease beta and rc tags', () => {
  assert.match(releaseWorkflow, /tags:\s*\n\s*-\s*'v\*'/);
  assert.match(releaseWorkflow, /tags:\s*\n(?:.*\n)*\s*-\s*'!v\*-beta\.\*'/);
  assert.match(releaseWorkflow, /tags:\s*\n(?:.*\n)*\s*-\s*'!v\*-rc\.\*'/);
});

test('stable release tags publish docs and prereleases do not update stable docs', () => {
  assert.match(docsPagesWorkflow, /tags:\s*\n\s*-\s*'v\*'/);
  assert.match(docsPagesWorkflow, /github\.ref_name/);
  assert.match(docsPagesWorkflow, /\^v\[0-9\]\+\\\.\[0-9\]\+\\\.\[0-9\]\+\$/);
  assert.match(docsPagesWorkflow, /bun run docs:build:versioned/);
  assert.doesNotMatch(docsPagesWorkflow, /beta/);
});

test('publish release forces an existing draft tag release to become public', () => {
  assert.ok(releaseWorkflow.includes('--draft=false'));
});

test('release workflow verifies a committed changelog section before publish', () => {
  assert.match(releaseWorkflow, /bun run changelog:check/);
});

test('release workflow guards against pending changelog fragments instead of auto-building them', () => {
  assert.match(releaseWorkflow, /Guard against pending changelog fragments/);
  assert.match(releaseWorkflow, /::error::Pending changelog fragments detected/);
  assert.match(releaseWorkflow, /Run 'bun run changelog:build --version/);
  assert.doesNotMatch(releaseWorkflow, /Build changelog artifacts for release/);
  assert.doesNotMatch(releaseWorkflow, /bun run changelog:build --version "\$\{\{ steps\.version/);
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
    /files=\(release\/\*\.AppImage release\/\*\.dmg release\/\*\.exe release\/\*\.zip release\/\*\.tar\.gz release\/latest\*\.yml release\/\*\.blockmap dist\/launcher\/subminer\)/,
  );
  assert.match(
    releaseWorkflow,
    /artifacts=\([\s\S]*release\/\*\.exe[\s\S]*release\/latest\*\.yml[\s\S]*release\/\*\.blockmap[\s\S]*release\/SHA256SUMS\.txt[\s\S]*\)/,
  );
});

test('release workflow uploads updater metadata without builder debug YAML files', () => {
  assert.match(releaseWorkflow, /release\/latest\*\.yml/);
  assert.doesNotMatch(releaseWorkflow, /release\/\*\.yml/);
});

test('release package metadata enables GitHub updater metadata without builder uploads', () => {
  assert.equal(packageJson.build?.publish?.[0]?.provider, 'github');
  assert.equal(packageJson.build?.publish?.[0]?.owner, 'ksyasuda');
  assert.equal(packageJson.build?.publish?.[0]?.repo, 'SubMiner');
  assert.equal(packageJson.build?.electronUpdaterCompatibility, '>=2.16');
});

test('release workflow writes checksum entries using release asset basenames', () => {
  assert.match(releaseWorkflow, /: > release\/SHA256SUMS\.txt/);
  assert.match(releaseWorkflow, /for file in "\$\{files\[@\]\}"; do/);
  assert.match(releaseWorkflow, /\$\{file##\*\/\}/);
  assert.doesNotMatch(releaseWorkflow, /sha256sum "\$\{files\[@\]\}" > release\/SHA256SUMS\.txt/);
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

test('release packaging stages generated launcher as an app resource', () => {
  assert.ok(
    packageJson.build?.extraResources?.some(
      (resource) =>
        resource.from === 'dist/launcher/subminer' && resource.to === 'launcher/subminer',
    ),
  );
  assert.match(packageJson.scripts.build ?? '', /bun run build:launcher/);
  assert.match(packageJson.scripts['build:launcher'] ?? '', /--banner='#!\/usr\/bin\/env bun'/);
});

test('release packaging does not reference removed Windows window helper script', () => {
  assert.ok(
    !(packageJson.build?.extraResources ?? []).some((resource) =>
      [resource.from, resource.to].some((value) => value?.includes('get-mpv-window-windows.ps1')),
    ),
  );
});

test('Makefile clean preserves committed prerelease notes', () => {
  assert.match(makefile, /PRERELEASE_NOTES_BACKUP/);
  assert.match(makefile, /release\/prerelease-notes\.md/);
  assert.doesNotMatch(makefile, /clean:[\s\S]*@rm -rf dist release\n/);
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

test('release artifact names are distinct before upload', () => {
  assert.equal(packageJson.build?.mac?.artifactName, 'SubMiner-${version}-mac.${ext}');
  assert.equal(packageJson.build?.dmg?.artifactName, 'SubMiner-${version}.${ext}');
  assert.equal(packageJson.build?.win?.artifactName, 'SubMiner-${version}-win.${ext}');
  assert.equal(packageJson.build?.nsis?.artifactName, 'SubMiner-${version}.${ext}');
  assert.doesNotMatch(releaseWorkflow, /Rename Windows ZIP artifacts/);
  assert.doesNotMatch(releaseWorkflow, /Rename-Item[\s\S]*-win\.zip/);
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

test('Makefile does not expose the legacy global mpv plugin installer', () => {
  assert.match(
    makefile,
    /windows\) printf '%s\\n' "\[INFO\] Windows builds run via: bun run build:win" ;;/,
  );
  assert.doesNotMatch(makefile, /^\s*install-plugin:/m);
  assert.doesNotMatch(makefile, /\binstall-plugin\b/);
  assert.doesNotMatch(makefile, /configure-plugin-binary-path\.mjs/);
});

test('Makefile uninstall targets remove bundled runtime plugin app-data copies', () => {
  assert.match(makefile, /uninstall-linux:[\s\S]*@rm -rf "\$\(LINUX_DATA_DIR\)\/plugin\/subminer"/);
  assert.match(makefile, /uninstall-macos:[\s\S]*@rm -rf "\$\(MACOS_DATA_DIR\)\/plugin\/subminer"/);
  assert.match(makefile, /Removed:[\s\S]*\$\(LINUX_DATA_DIR\)\/plugin\/subminer/);
  assert.match(makefile, /Removed:[\s\S]*\$\(MACOS_DATA_DIR\)\/plugin\/subminer/);
});
