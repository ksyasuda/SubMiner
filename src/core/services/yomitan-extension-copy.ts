import { createHash } from 'node:crypto';
import * as fs from 'fs';
import * as path from 'path';

const YOMITAN_SYNC_SCRIPT_PATHS = [
  path.join('js', 'app', 'popup.js'),
  path.join('js', 'display', 'popup-main.js'),
  path.join('js', 'display', 'display.js'),
  path.join('js', 'display', 'display-audio.js'),
];

function readManifestVersion(manifestPath: string): string | null {
  try {
    const parsed = JSON.parse(fs.readFileSync(manifestPath, 'utf-8')) as { version?: unknown };
    return typeof parsed.version === 'string' ? parsed.version : null;
  } catch {
    return null;
  }
}

export function hashDirectoryContents(dirPath: string): string | null {
  try {
    if (!fs.existsSync(dirPath) || !fs.statSync(dirPath).isDirectory()) {
      return null;
    }

    const hash = createHash('sha256');
    const queue = [''];
    while (queue.length > 0) {
      const relativeDir = queue.shift()!;
      const absoluteDir = path.join(dirPath, relativeDir);
      const entries = fs.readdirSync(absoluteDir, { withFileTypes: true });
      entries.sort((a, b) => a.name.localeCompare(b.name));

      for (const entry of entries) {
        const relativePath = path.join(relativeDir, entry.name);
        const normalizedRelativePath = relativePath.split(path.sep).join('/');
        hash.update(normalizedRelativePath);
        if (entry.isDirectory()) {
          queue.push(relativePath);
          continue;
        }
        if (!entry.isFile()) {
          continue;
        }
        hash.update(fs.readFileSync(path.join(dirPath, relativePath)));
      }
    }

    return hash.digest('hex');
  } catch {
    return null;
  }
}

function areFilesEqual(sourcePath: string, targetPath: string): boolean {
  if (!fs.existsSync(sourcePath) || !fs.existsSync(targetPath)) return false;
  try {
    return fs.readFileSync(sourcePath).equals(fs.readFileSync(targetPath));
  } catch {
    return false;
  }
}

export function shouldCopyYomitanExtension(sourceDir: string, targetDir: string): boolean {
  if (!fs.existsSync(targetDir)) {
    return true;
  }

  const sourceManifest = path.join(sourceDir, 'manifest.json');
  const targetManifest = path.join(targetDir, 'manifest.json');
  if (!fs.existsSync(sourceManifest) || !fs.existsSync(targetManifest)) {
    return true;
  }

  const sourceVersion = readManifestVersion(sourceManifest);
  const targetVersion = readManifestVersion(targetManifest);
  if (sourceVersion === null || targetVersion === null || sourceVersion !== targetVersion) {
    return true;
  }

  for (const relativePath of YOMITAN_SYNC_SCRIPT_PATHS) {
    if (!areFilesEqual(path.join(sourceDir, relativePath), path.join(targetDir, relativePath))) {
      return true;
    }
  }

  const sourceHash = hashDirectoryContents(sourceDir);
  const targetHash = hashDirectoryContents(targetDir);
  return sourceHash === null || targetHash === null || sourceHash !== targetHash;
}

export function ensureExtensionCopy(sourceDir: string, userDataPath: string): {
  targetDir: string;
  copied: boolean;
} {
  if (process.platform === 'win32') {
    return { targetDir: sourceDir, copied: false };
  }

  const extensionsRoot = path.join(userDataPath, 'extensions');
  const targetDir = path.join(extensionsRoot, 'yomitan');

  let shouldCopy = !fs.existsSync(targetDir);
  if (!shouldCopy) {
    shouldCopy = hashDirectoryContents(sourceDir) !== hashDirectoryContents(targetDir);
  }

  if (shouldCopy) {
    fs.mkdirSync(extensionsRoot, { recursive: true });
    fs.rmSync(targetDir, { recursive: true, force: true });
    fs.cpSync(sourceDir, targetDir, { recursive: true });
  }

  return { targetDir, copied: shouldCopy };
}
