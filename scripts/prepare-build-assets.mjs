import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, '..');
const rendererSourceDir = path.join(repoRoot, 'src', 'renderer');
const rendererOutputDir = path.join(repoRoot, 'dist', 'renderer');
const settingsSourceDir = path.join(repoRoot, 'src', 'settings');
const settingsOutputDir = path.join(repoRoot, 'dist', 'settings');
const syncUiSourceDir = path.join(repoRoot, 'src', 'syncui');
const syncUiOutputDir = path.join(repoRoot, 'dist', 'syncui');
const animeUiSourceDir = path.join(repoRoot, 'src', 'animeui');
const animeUiOutputDir = path.join(repoRoot, 'dist', 'animeui');
const scriptsOutputDir = path.join(repoRoot, 'dist', 'scripts');
const macosHelperSourcePath = path.join(scriptDir, 'get-mpv-window-macos.swift');
const macosHelperBinaryPath = path.join(scriptsOutputDir, 'get-mpv-window-macos');
const macosHelperSourceCopyPath = path.join(scriptsOutputDir, 'get-mpv-window-macos.swift');

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function copyFile(sourcePath, outputPath) {
  ensureDir(path.dirname(outputPath));
  fs.copyFileSync(sourcePath, outputPath);
}

function copyAssets(sourceDir, outputDir, label, stylesheets = ['style.css']) {
  copyFile(path.join(sourceDir, 'index.html'), path.join(outputDir, 'index.html'));
  for (const stylesheet of stylesheets) {
    copyFile(path.join(sourceDir, stylesheet), path.join(outputDir, stylesheet));
  }
  fs.cpSync(path.join(rendererSourceDir, 'fonts'), path.join(outputDir, 'fonts'), {
    recursive: true,
    force: true,
  });
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

function copyAnimeUiAssets() {
  copyAssets(animeUiSourceDir, animeUiOutputDir, 'animeui', [
    'style.css',
    'detail.css',
    'panels.css',
  ]);
}

function fallbackToMacosSource() {
  copyFile(macosHelperSourcePath, macosHelperSourceCopyPath);
  process.stdout.write(`Staged macOS helper source fallback: ${macosHelperSourceCopyPath}\n`);
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

function buildMacosHelper() {
  if (shouldSkipMacosHelperBuild()) {
    process.stdout.write('Skipping macOS helper build (SUBMINER_SKIP_MACOS_HELPER_BUILD=1)\n');
    fallbackToMacosSource();
    return;
  }

  if (process.platform !== 'darwin') {
    process.stdout.write('Skipping macOS helper build (not on macOS)\n');
    fallbackToMacosSource();
    return;
  }

  ensureDir(scriptsOutputDir);

  try {
    execFileSync(
      'swiftc',
      ['-O', '-target', macosHelperTarget(), macosHelperSourcePath, '-o', macosHelperBinaryPath],
      {
        stdio: 'inherit',
      },
    );
    fs.chmodSync(macosHelperBinaryPath, 0o755);
    process.stdout.write(`Built macOS helper: ${macosHelperBinaryPath}\n`);
  } catch (error) {
    process.stdout.write('Failed to compile macOS helper; using source fallback.\n');
    fallbackToMacosSource();
    if (error instanceof Error) {
      process.stderr.write(`${error.message}\n`);
    }
  }
}

function main() {
  copyRendererAssets();
  copySettingsAssets();
  copySyncUiAssets();
  copyAnimeUiAssets();
  buildMacosHelper();
}

main();
