import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { resolveLogBaseDir } from '../../shared/log-files';
import { writeStoredZip } from '../../shared/stored-zip';
import { redactLogExportText } from './log-redaction';

type LogCandidate = {
  path: string;
  name: string;
  kind: string;
  mtimeMs: number;
  mtimeDateKey: string;
  fileDateKey: string | null;
  fileWeekKey: string | null;
};

export type ExportLogsResult = {
  zipPath: string;
  exportedFiles: string[];
  mode: 'current-day' | 'most-recent';
};

export type ExportLogsOptions = {
  platform?: NodeJS.Platform;
  homeDir?: string;
  appDataDir?: string;
  logsDir?: string;
  outputDir?: string;
  now?: Date;
};

function pad(value: number): string {
  return String(value).padStart(2, '0');
}

function localDateKey(date: Date): string {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function localWeekKey(date: Date): string {
  const startOfYear = new Date(date.getFullYear(), 0, 1);
  const dayOfYear =
    Math.floor((date.getTime() - startOfYear.getTime()) / (24 * 60 * 60 * 1000)) + 1;
  return `${date.getFullYear()}-W${pad(Math.max(1, Math.ceil(dayOfYear / 7)))}`;
}

function filenameDateKey(fileName: string): string | null {
  return fileName.match(/\d{4}-\d{2}-\d{2}/)?.[0] ?? null;
}

function filenameWeekKey(fileName: string): string | null {
  return fileName.match(/\d{4}-W\d{2}/)?.[0] ?? null;
}

function fileKind(fileName: string): string {
  const match = fileName.match(/^([A-Za-z0-9_-]+)-/);
  return match?.[1] ?? fileName;
}

function zipTimestamp(date: Date): string {
  return `${localDateKey(date)}-${pad(date.getHours())}${pad(date.getMinutes())}${pad(
    date.getSeconds(),
  )}`;
}

function resolveLogsDir(options: ExportLogsOptions): string {
  if (options.logsDir) return options.logsDir;
  return path.join(
    resolveLogBaseDir({
      platform: options.platform,
      homeDir: options.homeDir,
      appDataDir: options.appDataDir,
    }),
    'logs',
  );
}

function buildCandidate(logsDir: string, entry: string): LogCandidate | null {
  if (!entry.endsWith('.log')) return null;

  const candidatePath = path.join(logsDir, entry);
  let stats: fs.Stats;
  try {
    stats = fs.statSync(candidatePath);
  } catch {
    return null;
  }
  if (!stats.isFile()) return null;

  return {
    path: candidatePath,
    name: entry,
    kind: fileKind(entry),
    mtimeMs: stats.mtimeMs,
    mtimeDateKey: localDateKey(stats.mtime),
    fileDateKey: filenameDateKey(entry),
    fileWeekKey: filenameWeekKey(entry),
  };
}

function listLogCandidates(logsDir: string): LogCandidate[] {
  let entries: string[];
  try {
    entries = fs.readdirSync(logsDir);
  } catch {
    return [];
  }

  return entries
    .map((entry) => buildCandidate(logsDir, entry))
    .filter((entry): entry is LogCandidate => entry !== null)
    .sort((left, right) => left.name.localeCompare(right.name));
}

function selectMostRecentPerKind(candidates: LogCandidate[]): LogCandidate[] {
  const byKind = new Map<string, LogCandidate>();
  for (const candidate of [...candidates].sort(
    (left, right) => candidateFreshnessMs(right) - candidateFreshnessMs(left),
  )) {
    if (!byKind.has(candidate.kind)) {
      byKind.set(candidate.kind, candidate);
    }
  }
  return [...byKind.values()].sort((left, right) => left.name.localeCompare(right.name));
}

function candidateFreshnessMs(candidate: LogCandidate): number {
  if (candidate.fileDateKey) {
    return Date.parse(`${candidate.fileDateKey}T23:59:59.999Z`);
  }
  if (candidate.fileWeekKey) {
    const match = candidate.fileWeekKey.match(/^(\d{4})-W(\d{2})$/);
    if (match) {
      const year = Number(match[1]);
      const week = Number(match[2]);
      return Date.UTC(year, 0, week * 7, 23, 59, 59, 999);
    }
  }
  return candidate.mtimeMs;
}

function selectLogCandidates(
  candidates: LogCandidate[],
  now: Date,
): { mode: ExportLogsResult['mode']; selected: LogCandidate[] } {
  const today = localDateKey(now);
  const currentDated = candidates.filter((candidate) => candidate.fileDateKey === today);
  if (currentDated.length > 0) {
    return { mode: 'current-day', selected: currentDated };
  }

  const currentWeek = localWeekKey(now);
  const currentWeekly = candidates.filter((candidate) => candidate.fileWeekKey === currentWeek);
  if (currentWeekly.length > 0) {
    return { mode: 'current-day', selected: currentWeekly };
  }

  const currentUndated = candidates.filter(
    (candidate) => candidate.fileDateKey === null && candidate.mtimeDateKey === today,
  );
  if (currentUndated.length > 0) {
    return { mode: 'current-day', selected: currentUndated };
  }
  return { mode: 'most-recent', selected: selectMostRecentPerKind(candidates) };
}

export function maskUsernamesInLogText(text: string): string {
  return redactLogExportText(text);
}

export function exportLogsArchive(options: ExportLogsOptions = {}): ExportLogsResult {
  const now = options.now ?? new Date();
  const logsDir = resolveLogsDir(options);
  const outputDir = options.outputDir ?? logsDir;
  const candidates = listLogCandidates(logsDir);
  const { mode, selected } = selectLogCandidates(candidates, now);

  if (selected.length === 0) {
    throw new Error(`No SubMiner log files found in ${logsDir}`);
  }

  fs.mkdirSync(outputDir, { recursive: true });
  const zipPath = path.join(outputDir, `subminer-logs-${zipTimestamp(now)}.zip`);

  writeStoredZip(
    zipPath,
    selected.map((candidate) => ({
      name: `logs/${candidate.name}`,
      data: Buffer.from(maskUsernamesInLogText(fs.readFileSync(candidate.path, 'utf8')), 'utf8'),
    })),
  );

  return {
    zipPath,
    exportedFiles: selected.map((candidate) => candidate.path),
    mode,
  };
}

export function exportLogsArchiveForCurrentUser(now: Date = new Date()): ExportLogsResult {
  return exportLogsArchive({
    platform: process.platform,
    homeDir: os.homedir(),
    appDataDir: process.env.APPDATA,
    now,
  });
}
