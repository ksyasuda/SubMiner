import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Database } from 'bun:sqlite';
import { SCHEMA_VERSION } from '../../src/core/services/immersion-tracker/types.js';
import { withReadonlyWalRetry } from '../history-db.js';
import { resolveConfigDir } from '../../src/config/path-resolution.js';

export { SCHEMA_VERSION };

export interface SyncMergeSummary {
  sessionsMerged: number;
  sessionsAlreadyPresent: number;
  activeSessionsSkipped: number;
  animeAdded: number;
  videosAdded: number;
  wordsAdded: number;
  kanjiAdded: number;
  subtitleLinesAdded: number;
  telemetryRowsAdded: number;
  eventsAdded: number;
  excludedWordsAdded: number;
  dailyRollupsCopied: number;
  monthlyRollupsCopied: number;
  rollupGroupsRecomputed: number;
}

export function createEmptyMergeSummary(): SyncMergeSummary {
  return {
    sessionsMerged: 0,
    sessionsAlreadyPresent: 0,
    activeSessionsSkipped: 0,
    animeAdded: 0,
    videosAdded: 0,
    wordsAdded: 0,
    kanjiAdded: 0,
    subtitleLinesAdded: 0,
    telemetryRowsAdded: 0,
    eventsAdded: 0,
    excludedWordsAdded: 0,
    dailyRollupsCopied: 0,
    monthlyRollupsCopied: 0,
    rollupGroupsRecomputed: 0,
  };
}

export function nowDbTimestamp(): string {
  return String(Date.now());
}

export function tableExists(db: Database, tableName: string): boolean {
  return Boolean(
    db.query(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(tableName),
  );
}

export function readSchemaVersion(db: Database): number | null {
  if (!tableExists(db, 'imm_schema_version')) return null;
  const row = db
    .query<{ schema_version: number }>(
      'SELECT MAX(schema_version) AS schema_version FROM imm_schema_version',
    )
    .get();
  return typeof row?.schema_version === 'number' ? row.schema_version : null;
}

export function assertMergeableSchema(db: Database, label: string): void {
  const version = readSchemaVersion(db);
  if (version === null) {
    throw new Error(
      `${label} database has no schema version. Run SubMiner once on that machine so the stats database is initialized.`,
    );
  }
  if (version !== SCHEMA_VERSION) {
    throw new Error(
      `${label} database is at schema version ${version} but this launcher expects ${SCHEMA_VERSION}. Update SubMiner on both machines to the same version and run each app once before syncing.`,
    );
  }
  for (const table of ['imm_sessions', 'imm_videos', 'imm_lifetime_global']) {
    if (!tableExists(db, table)) {
      throw new Error(`${label} database is missing table ${table}; cannot sync.`);
    }
  }
}

export function insertRow(
  db: Database,
  table: string,
  columns: readonly string[],
  values: unknown[],
): number {
  const sql = `INSERT INTO ${table} (${columns.join(', ')}) VALUES (${columns.map(() => '?').join(', ')})`;
  // db.query() caches the prepared statement per SQL string; this runs once
  // per copied row, so re-preparing via db.prepare() would dominate merge time.
  const result = db.query(sql).run(...values);
  return Number(result.lastInsertRowid);
}

export function createDbSnapshot(dbPath: string, outPath: string): void {
  if (!fs.existsSync(dbPath)) {
    throw new Error(`Stats database not found: ${dbPath}`);
  }
  fs.rmSync(outPath, { force: true });
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  withReadonlyWalRetry(dbPath, (options) => {
    const db = new Database(dbPath, options);
    try {
      assertMergeableSchema(db, 'Local');
      db.prepare('VACUUM INTO ?').run(outPath);
    } finally {
      db.close();
    }
  });
}

interface DaemonStateFile {
  pid?: unknown;
}

function isProcessAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    // EPERM means the process exists but we can't signal it → still alive.
    // Only ESRCH (no such process) means it's actually gone.
    return (error as NodeJS.ErrnoException)?.code === 'EPERM';
  }
}

function statsDaemonStateCandidates(dbPath: string): string[] {
  const homeDir = os.homedir();
  const candidates = new Set<string>([path.join(path.dirname(dbPath), 'stats-daemon.json')]);
  const configDir = resolveConfigDir({
    platform: process.platform,
    appDataDir: process.env.APPDATA,
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir,
    existsSync: fs.existsSync,
  });
  candidates.add(path.join(configDir, 'stats-daemon.json'));
  if (process.platform === 'darwin') {
    candidates.add(path.join(homeDir, 'Library', 'Application Support', 'SubMiner', 'stats-daemon.json'));
  }
  return [...candidates];
}

/**
 * Best-effort guard against merging while a SubMiner process holds the
 * tracker's write queue in memory. Detects the background stats daemon via
 * its pid state file; the interactive app is caught by the mpv-socket check
 * in the sync command.
 */
export function findLiveStatsDaemonPid(dbPath: string): number | null {
  for (const statePath of statsDaemonStateCandidates(dbPath)) {
    let raw: string;
    try {
      raw = fs.readFileSync(statePath, 'utf8');
    } catch {
      continue;
    }
    try {
      const parsed = JSON.parse(raw) as DaemonStateFile;
      const pid = typeof parsed.pid === 'number' && Number.isInteger(parsed.pid) ? parsed.pid : 0;
      if (pid > 0 && isProcessAlive(pid)) {
        return pid;
      }
    } catch {
      continue;
    }
  }
  return null;
}
