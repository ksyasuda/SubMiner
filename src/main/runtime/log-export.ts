import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { resolveLogBaseDir } from '../../shared/log-files';
import { writeStoredZip } from '../../shared/stored-zip';

type LogCandidate = {
  path: string;
  name: string;
  kind: string;
  mtimeMs: number;
  dateKeys: Set<string>;
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

const REDACTED_USER = '<user>';

function pad(value: number): string {
  return String(value).padStart(2, '0');
}

function localDateKey(date: Date): string {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function filenameDateKey(fileName: string): string | null {
  return fileName.match(/\d{4}-\d{2}-\d{2}/)?.[0] ?? null;
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

  const dateKeys = new Set<string>([localDateKey(stats.mtime)]);
  const fromName = filenameDateKey(entry);
  if (fromName) dateKeys.add(fromName);

  return {
    path: candidatePath,
    name: entry,
    kind: fileKind(entry),
    mtimeMs: stats.mtimeMs,
    dateKeys,
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
  for (const candidate of [...candidates].sort((left, right) => right.mtimeMs - left.mtimeMs)) {
    if (!byKind.has(candidate.kind)) {
      byKind.set(candidate.kind, candidate);
    }
  }
  return [...byKind.values()].sort((left, right) => left.name.localeCompare(right.name));
}

function selectLogCandidates(
  candidates: LogCandidate[],
  now: Date,
): { mode: ExportLogsResult['mode']; selected: LogCandidate[] } {
  const today = localDateKey(now);
  const currentDay = candidates.filter((candidate) => candidate.dateKeys.has(today));
  if (currentDay.length > 0) {
    return { mode: 'current-day', selected: currentDay };
  }
  return { mode: 'most-recent', selected: selectMostRecentPerKind(candidates) };
}

export function maskUsernamesInLogText(text: string): string {
  return text
    .replace(/(\/(?:home|Users)\/)([^/\r\n]+)(?=\/|$)/g, `$1${REDACTED_USER}`)
    .replace(/([A-Za-z]:[\\/]+Users[\\/]+)([^\\/:\r\n]+)(?=[\\/]|$)/g, `$1${REDACTED_USER}`);
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
