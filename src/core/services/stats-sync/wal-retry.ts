import fs from 'node:fs';
import type { SyncDbOpenOptions } from './driver';

/**
 * Opening a WAL-mode SQLite database strictly read-only fails when the -shm
 * file is missing or stale (the reader must be able to create it). Retry such
 * failures with a read-write handle; the caller still only issues reads.
 */
export function withReadonlyWalRetry<T>(
  dbPath: string,
  query: (options: SyncDbOpenOptions) => T,
): T {
  try {
    return query({ readonly: true });
  } catch (error) {
    if (!isReadonlyWalRetryError(error, dbPath)) throw error;
    return query({ readwrite: true, create: false });
  }
}

export function isReadonlyWalRetryError(error: unknown, dbPath: string): boolean {
  if (!isWalModeSqliteDatabase(dbPath)) return false;
  const code =
    typeof error === 'object' && error !== null && 'code' in error
      ? String((error as { code?: unknown }).code ?? '')
      : '';
  const message = error instanceof Error ? error.message : String(error);
  const text = `${code} ${message}`.toLowerCase();
  return (
    text.includes('readonly') ||
    text.includes('read-only') ||
    text.includes('attempt to write a readonly database') ||
    text.includes('sqlite_cantopen') ||
    text.includes('unable to open database file')
  );
}

function isWalModeSqliteDatabase(dbPath: string): boolean {
  const header = Buffer.alloc(20);
  let fd: number | null = null;
  try {
    fd = fs.openSync(dbPath, 'r');
    if (fs.readSync(fd, header, 0, header.length, 0) < header.length) return false;
  } catch {
    return false;
  } finally {
    if (fd !== null) fs.closeSync(fd);
  }
  return header.subarray(0, 16).toString('ascii') === 'SQLite format 3\0' && header[18] === 2;
}
