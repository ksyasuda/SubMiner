import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, '..');
const rendererSourceDir = path.join(repoRoot, 'src', 'renderer');
const rendererOutputDir = path.join(repoRoot, 'dist', 'renderer');
const settingsSourceDir = path.join(repoRoot, 'src', 'settings');
const settingsOutputDir = path.join(repoRoot, 'dist', 'settings');
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

function copyRendererAssets() {
  copyFile(path.join(rendererSourceDir, 'index.html'), path.join(rendererOutputDir, 'index.html'));
  copyFile(path.join(rendererSourceDir, 'style.css'), path.join(rendererOutputDir, 'style.css'));
  fs.cpSync(path.join(rendererSourceDir, 'fonts'), path.join(rendererOutputDir, 'fonts'), {
    recursive: true,
    force: true,
  });
  process.stdout.write(`Staged renderer assets in ${rendererOutputDir}\n`);
}

function copySettingsAssets() {
  copyFile(path.join(settingsSourceDir, 'index.html'), path.join(settingsOutputDir, 'index.html'));
  copyFile(path.join(settingsSourceDir, 'style.css'), path.join(settingsOutputDir, 'style.css'));
  fs.cpSync(path.join(rendererSourceDir, 'fonts'), path.join(settingsOutputDir, 'fonts'), {
    recursive: true,
    force: true,
  });
  process.stdout.write(`Staged settings assets in ${settingsOutputDir}\n`);
}

function fallbackToMacosSource() {
  copyFile(macosHelperSourcePath, macosHelperSourceCopyPath);
  process.stdout.write(`Staged macOS helper source fallback: ${macosHelperSourceCopyPath}\n`);
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

  try {
    execFileSync('swiftc', ['-O', macosHelperSourcePath, '-o', macosHelperBinaryPath], {
      stdio: 'inherit',
    });
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
  buildMacosHelper();
}

main();
