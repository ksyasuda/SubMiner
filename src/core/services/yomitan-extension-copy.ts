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

  return false;
}
