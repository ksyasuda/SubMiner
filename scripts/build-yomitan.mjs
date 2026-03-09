import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(dirname, '..');
const submoduleDir = path.join(repoRoot, 'vendor', 'subminer-yomitan');
const submodulePackagePath = path.join(submoduleDir, 'package.json');
const submodulePackageLockPath = path.join(submoduleDir, 'package-lock.json');
const buildOutputDir = path.join(repoRoot, 'build', 'yomitan');
const stampPath = path.join(buildOutputDir, '.subminer-build.json');
const zipPath = path.join(submoduleDir, 'builds', 'yomitan-chrome.zip');
const bunCommand = process.versions.bun ? process.execPath : 'bun';
const dependencyStampPath = path.join(submoduleDir, 'node_modules', '.subminer-package-lock-hash');

function run(command, args, cwd) {
  execFileSync(command, args, { cwd, stdio: 'inherit' });
}

function escapePowerShellString(value) {
  return value.replaceAll("'", "''");
}

function readCommand(command, args, cwd) {
  return execFileSync(command, args, { cwd, encoding: 'utf8' }).trim();
}

function readStamp() {
  try {
    return JSON.parse(fs.readFileSync(stampPath, 'utf8'));
  } catch {
    return null;
  }
}

function hashFile(filePath) {
  const hash = createHash('sha256');
  hash.update(fs.readFileSync(filePath));
  return hash.digest('hex');
}

function ensureSubmodulePresent() {
  if (!fs.existsSync(submodulePackagePath)) {
    throw new Error(
      'Missing vendor/subminer-yomitan submodule. Run `git submodule update --init --recursive`.',
    );
  }
}

function getSourceState() {
  const revision = readCommand('git', ['rev-parse', 'HEAD'], submoduleDir);
  const dirty = readCommand('git', ['status', '--short', '--untracked-files=no'], submoduleDir);
  return { revision, dirty };
}

function isBuildCurrent(force) {
  if (force) {
    return false;
  }
  if (!fs.existsSync(path.join(buildOutputDir, 'manifest.json'))) {
    return false;
  }

  const stamp = readStamp();
  if (!stamp) {
    return false;
  }

  const currentState = getSourceState();
  return stamp.revision === currentState.revision && stamp.dirty === currentState.dirty;
}

function ensureDependenciesInstalled() {
  const nodeModulesDir = path.join(submoduleDir, 'node_modules');
  const currentLockHash = hashFile(submodulePackageLockPath);
  let installedLockHash = '';
  try {
    installedLockHash = fs.readFileSync(dependencyStampPath, 'utf8').trim();
  } catch {}

  if (!fs.existsSync(nodeModulesDir) || installedLockHash !== currentLockHash) {
    run(bunCommand, ['install', '--no-save'], submoduleDir);
    fs.mkdirSync(nodeModulesDir, { recursive: true });
    fs.writeFileSync(dependencyStampPath, `${currentLockHash}\n`, 'utf8');
  }
}

function installAndBuild() {
  ensureDependenciesInstalled();
  run(bunCommand, ['./dev/bin/build.js', '--target', 'chrome'], submoduleDir);
}

function extractBuild() {
  if (!fs.existsSync(zipPath)) {
    throw new Error(`Expected Yomitan build artifact at ${zipPath}`);
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-yomitan-'));
  try {
    if (process.platform === 'win32') {
      run(
        'powershell.exe',
        [
          '-NoProfile',
          '-NonInteractive',
          '-ExecutionPolicy',
          'Bypass',
          '-Command',
          `Expand-Archive -LiteralPath '${escapePowerShellString(zipPath)}' -DestinationPath '${escapePowerShellString(tempDir)}' -Force`,
        ],
        repoRoot,
      );
    } else {
      run('unzip', ['-qo', zipPath, '-d', tempDir], repoRoot);
    }
    fs.rmSync(buildOutputDir, { recursive: true, force: true });
    fs.mkdirSync(path.dirname(buildOutputDir), { recursive: true });
    fs.cpSync(tempDir, buildOutputDir, { recursive: true });
    if (!fs.existsSync(path.join(buildOutputDir, 'manifest.json'))) {
      throw new Error(`Extracted Yomitan build missing manifest.json in ${buildOutputDir}`);
    }
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

function writeStamp() {
  const state = getSourceState();
  fs.writeFileSync(
    stampPath,
    `${JSON.stringify(
      {
        revision: state.revision,
        dirty: state.dirty,
        builtAt: new Date().toISOString(),
      },
      null,
      2,
    )}\n`,
    'utf8',
  );
}

function main() {
  const force = process.argv.includes('--force');
  ensureSubmodulePresent();

  if (isBuildCurrent(force)) {
    process.stdout.write(`Yomitan build current: ${buildOutputDir}\n`);
    return;
  }

  process.stdout.write('Building Yomitan Chrome artifact...\n');
  installAndBuild();
  extractBuild();
  writeStamp();
  process.stdout.write(`Yomitan extracted to ${buildOutputDir}\n`);
}

main();
