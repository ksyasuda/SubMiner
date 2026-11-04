import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { stageBundledIntegrationKeys, TMDB_API_KEY_ENV } from './bundled-integration-keys.mjs';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, '..');
const rendererSourceDir = path.join(repoRoot, 'src', 'renderer');
const rendererOutputDir = path.join(repoRoot, 'dist', 'renderer');
const settingsSourceDir = path.join(repoRoot, 'src', 'settings');
const settingsOutputDir = path.join(repoRoot, 'dist', 'settings');
const syncUiSourceDir = path.join(repoRoot, 'src', 'syncui');
const syncUiOutputDir = path.join(repoRoot, 'dist', 'syncui');
const scriptsOutputDir = path.join(repoRoot, 'dist', 'scripts');
// Swift helpers compiled on macOS and shipped in Resources/scripts.
const MACOS_HELPERS = ['get-mpv-window-macos', 'macos-controller-watch'];

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function copyFile(sourcePath, outputPath) {
  ensureDir(path.dirname(outputPath));
  fs.copyFileSync(sourcePath, outputPath);
}

function copyAssets(sourceDir, outputDir, label) {
  copyFile(path.join(sourceDir, 'index.html'), path.join(outputDir, 'index.html'));
  copyFile(path.join(sourceDir, 'style.css'), path.join(outputDir, 'style.css'));
  process.stdout.write(`Staged ${label} assets in ${outputDir}\n`);
}

function copyRendererAssets() {
  copyAssets(rendererSourceDir, rendererOutputDir, 'renderer');
}

function copySettingsAssets() {
  copyAssets(settingsSourceDir, settingsOutputDir, 'settings');
}

function copySyncUiAssets() {
  copyAssets(syncUiSourceDir, syncUiOutputDir, 'syncui');
}

function fallbackToMacosSource(helperName) {
  const sourceCopyPath = path.join(scriptsOutputDir, `${helperName}.swift`);
  copyFile(path.join(scriptDir, `${helperName}.swift`), sourceCopyPath);
  process.stdout.write(`Staged macOS helper source fallback: ${sourceCopyPath}\n`);
}

// Pin the minimum macOS to the app's own floor (Electron's `minos`). Without an
// explicit target, swiftc stamps the build machine's OS version as the binary's
// minimum and the helper fails to load on older systems (#213). The arch stays
// the host's, matching the single-arch app electron-builder packages here.
const MACOS_HELPER_DEPLOYMENT_TARGET = '12.0';

function macosHelperTarget() {
  return `${os.arch() === 'x64' ? 'x86_64' : 'arm64'}-apple-macos${MACOS_HELPER_DEPLOYMENT_TARGET}`;
}

function shouldSkipMacosHelperBuild() {
  return process.env.SUBMINER_SKIP_MACOS_HELPER_BUILD === '1';
}

function buildMacosHelper(helperName) {
  if (shouldSkipMacosHelperBuild()) {
    process.stdout.write('Skipping macOS helper build (SUBMINER_SKIP_MACOS_HELPER_BUILD=1)\n');
    fallbackToMacosSource(helperName);
    return;
  }

  if (process.platform !== 'darwin') {
    process.stdout.write('Skipping macOS helper build (not on macOS)\n');
    fallbackToMacosSource(helperName);
    return;
  }

  ensureDir(scriptsOutputDir);
  const sourcePath = path.join(scriptDir, `${helperName}.swift`);
  const binaryPath = path.join(scriptsOutputDir, helperName);

  try {
    execFileSync('swiftc', ['-O', '-target', macosHelperTarget(), sourcePath, '-o', binaryPath], {
      stdio: 'inherit',
    });
    fs.chmodSync(binaryPath, 0o755);
    process.stdout.write(`Built macOS helper: ${binaryPath}\n`);
  } catch (error) {
    process.stdout.write(`Failed to compile macOS helper ${helperName}; using source fallback.\n`);
    // A binary from an earlier build would otherwise pass the after-pack check and ship stale.
    fs.rmSync(binaryPath, { force: true });
    fallbackToMacosSource(helperName);
    if (error instanceof Error) {
      process.stderr.write(`${error.message}\n`);
    }
  }
}

// Only the key names are logged, never the values: CI masks secrets, but a
// stray echo would still leak them into local build logs.
function stageIntegrationKeys() {
  const staged = stageBundledIntegrationKeys(path.join(repoRoot, 'dist'));
  process.stdout.write(
    staged.length > 0
      ? `Staged bundled integration keys: ${staged.join(', ')}\n`
      : `No bundled integration keys (${TMDB_API_KEY_ENV} unset)\n`,
  );
}

function main() {
  fs.cpSync(path.join(rendererSourceDir, 'fonts'), path.join(repoRoot, 'dist', 'fonts'), {
    recursive: true,
    force: true,
  });
  copyRendererAssets();
  copySettingsAssets();
  copySyncUiAssets();
  MACOS_HELPERS.forEach(buildMacosHelper);
  stageIntegrationKeys();
}

main();
