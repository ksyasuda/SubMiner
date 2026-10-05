import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { spawn } from 'node:child_process';
import { availableParallelism } from 'node:os';
import { isAbsolute, join, relative, resolve } from 'node:path';
import { collectLaneFiles } from './test-lanes';

type LcovRecord = {
  sourceFile: string;
  functions: Map<string, number>;
  functionHits: Map<string, number>;
  lines: Map<number, number>;
  branches: Map<string, { line: number; block: string; branch: string; hits: number | null }>;
};

const repoRoot = resolve(__dirname, '..');

function parseCoverageDirArg(argv: string[]): string {
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === '--coverage-dir') {
      const next = argv[index + 1];
      if (typeof next !== 'string') {
        throw new Error('Missing value for --coverage-dir');
      }
      return next;
    }
  }
  return 'coverage';
}

export function resolveCoverageDir(repoRootDir: string, argv: string[]): string {
  const candidate = resolve(repoRootDir, parseCoverageDirArg(argv));
  const rel = relative(repoRootDir, candidate);
  if (isAbsolute(rel) || rel.startsWith('..')) {
    throw new Error(`--coverage-dir must be within repository: ${candidate}`);
  }
  return candidate;
}

function parseLcovReport(report: string): LcovRecord[] {
  const records: LcovRecord[] = [];
  let current: LcovRecord | null = null;

  const ensureCurrent = (): LcovRecord => {
    if (!current) {
      throw new Error('Malformed lcov report: record data before SF');
    }
    return current;
  };

  for (const rawLine of report.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    if (line.startsWith('TN:')) {
      continue;
    }
    if (line.startsWith('SF:')) {
      current = {
        sourceFile: line.slice(3),
        functions: new Map(),
        functionHits: new Map(),
        lines: new Map(),
        branches: new Map(),
      };
      continue;
    }
    if (line === 'end_of_record') {
      if (current) {
        records.push(current);
        current = null;
      }
      continue;
    }
    if (line.startsWith('FN:')) {
      const [lineNumber, ...nameParts] = line.slice(3).split(',');
      ensureCurrent().functions.set(nameParts.join(','), Number(lineNumber));
      continue;
    }
    if (line.startsWith('FNDA:')) {
      const [hits, ...nameParts] = line.slice(5).split(',');
      ensureCurrent().functionHits.set(nameParts.join(','), Number(hits));
      continue;
    }
    if (line.startsWith('DA:')) {
      const [lineNumber, hits] = line.slice(3).split(',');
      ensureCurrent().lines.set(Number(lineNumber), Number(hits));
      continue;
    }
    if (line.startsWith('BRDA:')) {
      const [lineNumber, block, branch, hits] = line.slice(5).split(',');
      if (
        lineNumber === undefined ||
        block === undefined ||
        branch === undefined ||
        hits === undefined
      ) {
        continue;
      }
      ensureCurrent().branches.set(`${lineNumber}:${block}:${branch}`, {
        line: Number(lineNumber),
        block,
        branch,
        hits: hits === '-' ? null : Number(hits),
      });
    }
  }

  if (current) {
    records.push(current);
  }

  return records;
}

export function mergeLcovReports(reports: string[]): string {
  const merged = new Map<string, LcovRecord>();

  for (const report of reports) {
    for (const record of parseLcovReport(report)) {
      let target = merged.get(record.sourceFile);
      if (!target) {
        target = {
          sourceFile: record.sourceFile,
          functions: new Map(),
          functionHits: new Map(),
          lines: new Map(),
          branches: new Map(),
        };
        merged.set(record.sourceFile, target);
      }

      for (const [name, line] of record.functions) {
        if (!target.functions.has(name)) {
          target.functions.set(name, line);
        }
      }

      for (const [name, hits] of record.functionHits) {
        target.functionHits.set(name, (target.functionHits.get(name) ?? 0) + hits);
      }

      for (const [lineNumber, hits] of record.lines) {
        target.lines.set(lineNumber, (target.lines.get(lineNumber) ?? 0) + hits);
      }

      for (const [branchKey, branchRecord] of record.branches) {
        const existing = target.branches.get(branchKey);
        if (!existing) {
          target.branches.set(branchKey, { ...branchRecord });
          continue;
        }
        if (branchRecord.hits === null) {
          continue;
        }
        existing.hits = (existing.hits ?? 0) + branchRecord.hits;
      }
    }
  }

  const chunks: string[] = [];
  for (const sourceFile of [...merged.keys()].sort()) {
    const record = merged.get(sourceFile)!;
    chunks.push(`SF:${record.sourceFile}`);

    const functions = [...record.functions.entries()].sort((a, b) =>
      a[1] === b[1] ? a[0].localeCompare(b[0]) : a[1] - b[1],
    );
    for (const [name, line] of functions) {
      chunks.push(`FN:${line},${name}`);
    }
    for (const [name] of functions) {
      chunks.push(`FNDA:${record.functionHits.get(name) ?? 0},${name}`);
    }
    chunks.push(`FNF:${functions.length}`);
    chunks.push(
      `FNH:${functions.filter(([name]) => (record.functionHits.get(name) ?? 0) > 0).length}`,
    );

    const branches = [...record.branches.values()].sort((a, b) =>
      a.line === b.line
        ? a.block === b.block
          ? a.branch.localeCompare(b.branch)
          : a.block.localeCompare(b.block)
        : a.line - b.line,
    );
    for (const branch of branches) {
      chunks.push(
        `BRDA:${branch.line},${branch.block},${branch.branch},${branch.hits === null ? '-' : branch.hits}`,
      );
    }
    chunks.push(`BRF:${branches.length}`);
    chunks.push(`BRH:${branches.filter((branch) => (branch.hits ?? 0) > 0).length}`);

    const lines = [...record.lines.entries()].sort((a, b) => a[0] - b[0]);
    for (const [lineNumber, hits] of lines) {
      chunks.push(`DA:${lineNumber},${hits}`);
    }
    chunks.push(`LF:${lines.length}`);
    chunks.push(`LH:${lines.filter(([, hits]) => hits > 0).length}`);
    chunks.push('end_of_record');
  }

  return chunks.length > 0 ? `${chunks.join('\n')}\n` : '';
}

type ShardResult = { status: number; output: string; lcov: string | null };

// Runs one test file under coverage into its own shard directory, buffering
// output so parallel shards do not interleave on the terminal.
function runShard(repoRootDir: string, file: string, shardDir: string): Promise<ShardResult> {
  return new Promise((resolveShard) => {
    const child = spawn(
      'bun',
      ['test', '--coverage', '--coverage-reporter=lcov', '--coverage-dir', shardDir, `./${file}`],
      { cwd: repoRootDir },
    );
    let output = '';
    child.stdout.on('data', (chunk) => (output += chunk));
    child.stderr.on('data', (chunk) => (output += chunk));
    child.on('error', (error) => resolveShard({ status: 1, output: String(error), lcov: null }));
    child.on('close', (code) => {
      const lcovPath = join(shardDir, 'lcov.info');
      const lcov = existsSync(lcovPath) ? readFileSync(lcovPath, 'utf8') : null;
      resolveShard({ status: code ?? 1, output, lcov });
    });
  });
}

function parseJobsArg(argv: string[]): number {
  const index = argv.indexOf('--jobs');
  if (index === -1) return availableParallelism();
  return Math.max(1, Number(argv[index + 1]) || 1);
}

// Shards run in parallel (one per CPU unless --jobs N). The first failing shard
// stops scheduling new shards and its exit status becomes the lane result.
export async function runCoverageLane(
  options: { repoRootDir?: string; argv?: string[]; quiet?: boolean } = {},
): Promise<number> {
  const repoRootDir = options.repoRootDir ?? repoRoot;
  const argv = options.argv ?? process.argv.slice(2);
  const laneName = argv[0];
  if (laneName === undefined) {
    process.stderr.write('Missing coverage lane name\n');
    return 1;
  }

  const coverageDir = resolveCoverageDir(repoRootDir, argv.slice(1));
  const shardRoot = join(coverageDir, '.shards');
  mkdirSync(coverageDir, { recursive: true });
  rmSync(shardRoot, { recursive: true, force: true });
  mkdirSync(shardRoot, { recursive: true });

  let files: string[];
  try {
    files = collectLaneFiles(repoRootDir, laneName);
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : error}\n`);
    return 1;
  }

  const reports: Array<string | null> = new Array(files.length).fill(null);
  const failures: Array<{ file: string; result: ShardResult }> = [];
  let nextIndex = 0;

  const worker = async (): Promise<void> => {
    while (failures.length === 0 && nextIndex < files.length) {
      const index = nextIndex++;
      const file = files[index]!;
      const shardDir = join(shardRoot, String(index + 1).padStart(3, '0'));
      const result = await runShard(repoRootDir, file, shardDir);
      if (result.status !== 0) {
        failures.push({ file, result });
        return;
      }
      if (result.lcov === null && !options.quiet) {
        process.stdout.write(`Skipping empty coverage shard for ${file}\n`);
      }
      reports[index] = result.lcov;
    }
  };

  try {
    const jobs = Math.min(parseJobsArg(argv), files.length);
    await Promise.all(Array.from({ length: jobs }, worker));

    const [failure] = failures;
    if (failure !== undefined) {
      const { file, result } = failure;
      if (!options.quiet) {
        process.stderr.write(`${result.output}\nCoverage shard failed: ${file}\n`);
      }
      return result.status;
    }

    const lcovPath = join(coverageDir, 'lcov.info');
    writeFileSync(lcovPath, mergeLcovReports(reports.filter((r) => r !== null)), 'utf8');
    process.stdout.write(`Merged LCOV written to ${relative(repoRootDir, lcovPath)}\n`);
    return 0;
  } finally {
    rmSync(shardRoot, { recursive: true, force: true });
  }
}

// @ts-ignore Bun entrypoint detection; TS config for scripts still targets CommonJS.
if (import.meta.main) {
  runCoverageLane().then((status) => process.exit(status));
}
