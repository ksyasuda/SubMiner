import fs from 'node:fs';
import path from 'node:path';

function normalizeCandidate(candidate) {
  if (typeof candidate !== 'string') return '';
  const trimmed = candidate.trim();
  return trimmed.length > 0 ? trimmed : '';
}

function fileExists(candidate) {
  try {
    return fs.statSync(candidate).isFile();
  } catch {
    return false;
  }
}

function unique(values) {
  return Array.from(new Set(values.filter((value) => value.length > 0)));
}

function findWindowsBinary(repoRoot) {
  const homeDir = process.env.HOME?.trim() || process.env.USERPROFILE?.trim() || '';
  const appDataDir = process.env.APPDATA?.trim() || '';
  const derivedLocalAppData =
    appDataDir && /[\\/]Roaming$/i.test(appDataDir)
      ? appDataDir.replace(/[\\/]Roaming$/i, `${path.sep}Local`)
      : '';
  const localAppData =
    process.env.LOCALAPPDATA?.trim() ||
    derivedLocalAppData ||
    (homeDir ? path.join(homeDir, 'AppData', 'Local') : '');
  const programFiles = process.env.ProgramFiles?.trim() || 'C:\\Program Files';
  const programFilesX86 = process.env['ProgramFiles(x86)']?.trim() || 'C:\\Program Files (x86)';

  const candidates = unique([
    normalizeCandidate(process.env.SUBMINER_BINARY_PATH),
    normalizeCandidate(process.env.SUBMINER_APPIMAGE_PATH),
    localAppData ? path.join(localAppData, 'Programs', 'SubMiner', 'SubMiner.exe') : '',
    path.join(programFiles, 'SubMiner', 'SubMiner.exe'),
    path.join(programFilesX86, 'SubMiner', 'SubMiner.exe'),
    'C:\\SubMiner\\SubMiner.exe',
    path.join(repoRoot, 'release', 'win-unpacked', 'SubMiner.exe'),
    path.join(repoRoot, 'release', 'SubMiner', 'SubMiner.exe'),
    path.join(repoRoot, 'release', 'SubMiner.exe'),
  ]);

  return candidates.find((candidate) => fileExists(candidate)) || '';
}

function rewriteBinaryPath(configPath, binaryPath) {
  const content = fs.readFileSync(configPath, 'utf8');
  const normalizedPath = binaryPath.replace(/\r?\n/g, ' ').trim();
  const updated = content.replace(/^binary_path=.*$/m, `binary_path=${normalizedPath}`);
  if (updated !== content) {
    fs.writeFileSync(configPath, updated, 'utf8');
  }
}

function rewriteSocketPath(configPath, socketPath) {
  const content = fs.readFileSync(configPath, 'utf8');
  const normalizedPath = socketPath.replace(/\r?\n/g, ' ').trim();
  const updated = content.replace(/^socket_path=.*$/m, `socket_path=${normalizedPath}`);
  if (updated !== content) {
    fs.writeFileSync(configPath, updated, 'utf8');
  }
}

const [, , configPathArg, repoRootArg, platformArg] = process.argv;
const configPath = normalizeCandidate(configPathArg);
const repoRoot = normalizeCandidate(repoRootArg) || process.cwd();
const platform = normalizeCandidate(platformArg) || process.platform;

if (!configPath) {
  console.error('[ERROR] Missing plugin config path');
  process.exit(1);
}

if (!fileExists(configPath)) {
  console.error(`[ERROR] Plugin config not found: ${configPath}`);
  process.exit(1);
}

if (platform !== 'win32') {
  console.log('[INFO] Skipping binary_path rewrite for non-Windows platform');
  process.exit(0);
}

const windowsSocketPath = '\\\\.\\pipe\\subminer-socket';
rewriteSocketPath(configPath, windowsSocketPath);

const binaryPath = findWindowsBinary(repoRoot);
if (!binaryPath) {
  console.warn(
    `[WARN] Configured plugin socket_path=${windowsSocketPath} but could not detect SubMiner.exe; set binary_path manually or provide SUBMINER_BINARY_PATH`,
  );
  process.exit(0);
}

rewriteBinaryPath(configPath, binaryPath);
console.log(`[INFO] Configured plugin socket_path=${windowsSocketPath} binary_path=${binaryPath}`);
