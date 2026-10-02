import fs from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';

export function ensureDir(dirPath: string): void {
  if (fs.existsSync(dirPath)) return;
  fs.mkdirSync(dirPath, { recursive: true });
}

export function ensureDirForFile(filePath: string): void {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

function syncDirectory(directoryPath: string): void {
  // Node cannot open Windows directory handles for fsync. The atomic rename still prevents
  // torn files there, while supported platforms also flush the directory entry below.
  if (process.platform === 'win32') return;

  const directory = fs.openSync(directoryPath, 'r');
  try {
    fs.fsyncSync(directory);
  } finally {
    fs.closeSync(directory);
  }
}

/** Writes and flushes a sibling temporary file before atomically replacing the target. */
export function writeTextFileAtomicallyDurable(filePath: string, content: string): void {
  const directoryPath = path.dirname(filePath);
  ensureDir(directoryPath);
  const temporaryPath = path.join(directoryPath, `.${path.basename(filePath)}.${randomUUID()}.tmp`);
  let temporaryFile: number | undefined;

  try {
    temporaryFile = fs.openSync(temporaryPath, 'wx', 0o600);
    fs.writeFileSync(temporaryFile, content, 'utf8');
    fs.fsyncSync(temporaryFile);
    fs.closeSync(temporaryFile);
    temporaryFile = undefined;
    fs.renameSync(temporaryPath, filePath);
    syncDirectory(directoryPath);
  } catch (error) {
    if (temporaryFile !== undefined) {
      try {
        fs.closeSync(temporaryFile);
      } catch {
        // Best effort cleanup preserves the original write failure.
      }
    }
    try {
      fs.rmSync(temporaryPath, { force: true });
    } catch {
      // The temporary file is disposable; the existing target stays untouched before rename.
    }
    throw error;
  }
}
